import * as React from 'react';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Toggle,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Label
} from '@fluentui/react';
import styles from './AuditItemForm.module.scss';
import {
  IAuditMasterItem,
  ICategoryItem,
  IISOClauseItem,
  ISegmentServiceItem,
  ICurrentUser,
  IValidationErrors,
  AuditStatus,
  AuditPriority,
  AuditType,
  AuditSource,
  VerificationResult,
  RiskRating
} from '../../models';
import { SharePointService, RoleService } from '../../services';

export interface IAuditItemFormProps {
  spService: SharePointService;
  roleService: RoleService;
  currentUser: ICurrentUser;
  editItem?: IAuditMasterItem;       // if provided, we're in edit mode
  onSaved: () => void;               // callback after successful save
  onCancel: () => void;
}

interface IFormState {
  formData: Partial<IAuditMasterItem>;
  categories: ICategoryItem[];
  isoClauses: IISOClauseItem[];
  segmentServices: ISegmentServiceItem[];
  errors: IValidationErrors;
  isSaving: boolean;
  isLoading: boolean;
  successMessage: string;
  errorMessage: string;
  isReadOnly: boolean;
}

/**
 * Full Audit Master item form – used for both Create and Edit.
 *
 * Sections:
 *  1. Basic Information (Title, AuditId, AuditYear, AuditType, Source)
 *  2. Classification (Category, ISO Clause, Segment/Service, Department, Location)
 *  3. People (PIC, Verifier, Auditor, Auditee)
 *  4. Finding Details (Finding, RootCause, CorrectiveAction, PreventiveAction)
 *  5. Dates & Status (Status, Priority, RiskRating, dates)
 *  6. Verification & Follow-up
 */
const AuditItemForm: React.FC<IAuditItemFormProps> = (props) => {
  const { spService, roleService, currentUser, editItem, onSaved, onCancel } = props;

  const isEditMode = !!editItem;

  const [state, setState] = React.useState<IFormState>({
    formData: editItem ? { ...editItem } : {
      Status: AuditStatus.Open,
      Priority: AuditPriority.Medium,
      VerificationResult: VerificationResult.NotVerified,
      RiskRating: RiskRating.Medium,
      ExtensionCount: 0,
      RequiresFollowUp: false
    },
    categories: [],
    isoClauses: [],
    segmentServices: [],
    errors: {},
    isSaving: false,
    isLoading: true,
    successMessage: '',
    errorMessage: '',
    isReadOnly: false
  });

  // ── Load reference data on mount ──────────────────────────────────────────
  React.useEffect(() => {
    const loadData = async (): Promise<void> => {
      try {
        const [categories, isoClauses, segmentServices] = await Promise.all([
          spService.getCategories(),
          spService.getISOClauses(),
          spService.getSegmentServices()
        ]);

        // Permission check
        let readOnly = false;
        if (isEditMode && editItem) {
          readOnly = !roleService.canUpdate(currentUser, editItem);
        } else {
          readOnly = !roleService.canCreate(currentUser);
        }

        setState(prev => ({
          ...prev,
          categories,
          isoClauses,
          segmentServices,
          isLoading: false,
          isReadOnly: readOnly
        }));
      } catch (err) {
        setState(prev => ({
          ...prev,
          isLoading: false,
          errorMessage: `Failed to load form data: ${err.message}`
        }));
      }
    };
    loadData();
  }, []);

  // ── Helpers ───────────────────────────────────────────────────────────────

  const updateField = (field: string, value: any): void => {
    setState(prev => ({
      ...prev,
      formData: { ...prev.formData, [field]: value },
      errors: { ...prev.errors, [field]: '' },
      successMessage: '',
      errorMessage: ''
    }));
  };

  const toDropdownOptions = (values: string[]): IDropdownOption[] =>
    values.map(v => ({ key: v, text: v }));

  const categoryOptions: IDropdownOption[] = state.categories.map(c => ({
    key: c.Id,
    text: c.Category
  }));

  const isoClauseOptions: IDropdownOption[] = state.isoClauses.map(c => ({
    key: c.Id,
    text: `${c.ISOClause}${c.Article ? ' – ' + c.Article : ''}`
  }));

  const segmentOptions: IDropdownOption[] = state.segmentServices.map(s => ({
    key: s.Id,
    text: `${s.Title} – ${s.Service}`
  }));

  // ── Validation ────────────────────────────────────────────────────────────

  const validate = (): IValidationErrors => {
    const errs: IValidationErrors = {};
    const fd = state.formData;

    if (!fd.Title || fd.Title.trim() === '') {
      errs.Title = 'Title is required.';
    }
    if (!fd.AuditId || fd.AuditId.trim() === '') {
      errs.AuditId = 'Audit ID is required.';
    }
    if (!fd.Status) {
      errs.Status = 'Status is required.';
    }
    if (!fd.AuditDate) {
      errs.AuditDate = 'Audit Date is required.';
    }
    if (fd.DueDate && fd.AuditDate && new Date(fd.DueDate) < new Date(fd.AuditDate)) {
      errs.DueDate = 'Due Date cannot be before Audit Date.';
    }
    if (!fd.Finding || fd.Finding.trim() === '') {
      errs.Finding = 'Finding is required.';
    }

    return errs;
  };

  // ── Submit ────────────────────────────────────────────────────────────────

  const handleSubmit = async (): Promise<void> => {
    const errs = validate();
    if (Object.keys(errs).length > 0) {
      setState(prev => ({ ...prev, errors: errs }));
      return;
    }

    setState(prev => ({ ...prev, isSaving: true, errorMessage: '' }));

    try {
      if (isEditMode && editItem) {
        await spService.updateAuditItem(editItem.Id, state.formData);
        setState(prev => ({
          ...prev,
          isSaving: false,
          successMessage: 'Audit item updated successfully!'
        }));
      } else {
        // Set created by user
        const payload = {
          ...state.formData,
          CreatedByUserId: currentUser.id
        };
        await spService.createAuditItem(payload);
        setState(prev => ({
          ...prev,
          isSaving: false,
          successMessage: 'Audit item created successfully!'
        }));
      }

      // Notify parent after short delay so user can see success message
      setTimeout(() => onSaved(), 1200);
    } catch (err) {
      setState(prev => ({
        ...prev,
        isSaving: false,
        errorMessage: `Save failed: ${err.message}`
      }));
    }
  };

  // ── Render ────────────────────────────────────────────────────────────────

  if (state.isLoading) {
    return (
      <div className={styles.formContainer}>
        <Spinner size={SpinnerSize.large} label="Loading form..." />
      </div>
    );
  }

  const fd = state.formData;
  const errs = state.errors;
  const ro = state.isReadOnly;

  return (
    <div className={styles.formContainer}>
      <div className={styles.formTitle}>
        {isEditMode ? `Edit Audit Item – ${fd.AuditId || ''}` : 'Create New Audit Item'}
      </div>
      <div className={styles.formSubtitle}>
        {isEditMode
          ? 'Modify the audit finding details below.'
          : 'Fill in the details to create a new audit finding.'}
      </div>

      {state.successMessage && (
        <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
          {state.successMessage}
        </MessageBar>
      )}
      {state.errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {state.errorMessage}
        </MessageBar>
      )}
      {ro && (
        <div className={styles.readOnlyBanner}>
          You do not have permission to edit this item. Displaying in read-only mode.
        </div>
      )}

      {/* ── Section 1: Basic Information ──────────────────────────────── */}
      <div className={styles.section}>
        <h4>Basic Information</h4>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Title"
              required
              disabled={ro}
              value={fd.Title || ''}
              onChange={(_, v) => updateField('Title', v)}
              errorMessage={errs.Title}
            />
          </div>
          <div>
            <TextField
              label="Audit ID"
              required
              disabled={ro}
              value={fd.AuditId || ''}
              onChange={(_, v) => updateField('AuditId', v)}
              errorMessage={errs.AuditId}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Audit Year"
              disabled={ro}
              value={fd.AuditYear || ''}
              onChange={(_, v) => updateField('AuditYear', v)}
              placeholder="e.g. 2026"
            />
          </div>
          <div>
            <Dropdown
              label="Audit Type"
              disabled={ro}
              selectedKey={fd.AuditType || undefined}
              options={toDropdownOptions(Object.values(AuditType))}
              onChange={(_, opt) => updateField('AuditType', opt?.key)}
            />
          </div>
          <div>
            <Dropdown
              label="Source"
              disabled={ro}
              selectedKey={fd.Source || undefined}
              options={toDropdownOptions(Object.values(AuditSource))}
              onChange={(_, opt) => updateField('Source', opt?.key)}
            />
          </div>
        </div>
        <div className={styles.fieldFull}>
          <TextField
            label="Description"
            multiline
            rows={3}
            disabled={ro}
            value={fd.Description || ''}
            onChange={(_, v) => updateField('Description', v)}
          />
        </div>
      </div>

      {/* ── Section 2: Classification ─────────────────────────────────── */}
      <div className={styles.section}>
        <h4>Classification</h4>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="Category"
              disabled={ro}
              selectedKey={fd.CategoryId || undefined}
              options={categoryOptions}
              onChange={(_, opt) => updateField('CategoryId', opt?.key)}
              placeholder="Select category"
            />
          </div>
          <div>
            <Dropdown
              label="ISO Clause"
              disabled={ro}
              selectedKey={fd.ISOClauseId || undefined}
              options={isoClauseOptions}
              onChange={(_, opt) => updateField('ISOClauseId', opt?.key)}
              placeholder="Select ISO clause"
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="Segment / Service"
              disabled={ro}
              selectedKey={fd.SegmentServiceId || undefined}
              options={segmentOptions}
              onChange={(_, opt) => updateField('SegmentServiceId', opt?.key)}
              placeholder="Select segment"
            />
          </div>
          <div>
            <TextField
              label="Department"
              disabled={ro}
              value={fd.Department || ''}
              onChange={(_, v) => updateField('Department', v)}
            />
          </div>
          <div>
            <TextField
              label="Location"
              disabled={ro}
              value={fd.Location || ''}
              onChange={(_, v) => updateField('Location', v)}
            />
          </div>
        </div>
      </div>

      {/* ── Section 3: People ─────────────────────────────────────────── */}
      <div className={styles.section}>
        <h4>People</h4>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="PIC (User ID)"
              disabled={ro}
              value={fd.PICId ? String(fd.PICId) : ''}
              onChange={(_, v) => updateField('PICId', v ? parseInt(v, 10) : undefined)}
              description="Enter the SharePoint User ID of the Person In Charge"
              type="number"
            />
          </div>
          <div>
            <TextField
              label="Verifier (User ID)"
              disabled={ro}
              value={fd.VerifierId ? String(fd.VerifierId) : ''}
              onChange={(_, v) => updateField('VerifierId', v ? parseInt(v, 10) : undefined)}
              description="Enter the SharePoint User ID of the Verifier"
              type="number"
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Auditor (User ID)"
              disabled={ro}
              value={fd.AuditorId ? String(fd.AuditorId) : ''}
              onChange={(_, v) => updateField('AuditorId', v ? parseInt(v, 10) : undefined)}
              type="number"
            />
          </div>
          <div>
            <TextField
              label="Auditee (User ID)"
              disabled={ro}
              value={fd.AuditeeId ? String(fd.AuditeeId) : ''}
              onChange={(_, v) => updateField('AuditeeId', v ? parseInt(v, 10) : undefined)}
              type="number"
            />
          </div>
        </div>
      </div>

      {/* ── Section 4: Finding Details ────────────────────────────────── */}
      <div className={styles.section}>
        <h4>Finding Details</h4>
        <div className={styles.fieldFull}>
          <TextField
            label="Finding"
            required
            multiline
            rows={3}
            disabled={ro}
            value={fd.Finding || ''}
            onChange={(_, v) => updateField('Finding', v)}
            errorMessage={errs.Finding}
          />
        </div>
        <div className={styles.fieldFull}>
          <TextField
            label="Root Cause"
            multiline
            rows={3}
            disabled={ro}
            value={fd.RootCause || ''}
            onChange={(_, v) => updateField('RootCause', v)}
          />
        </div>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Corrective Action"
              multiline
              rows={3}
              disabled={ro}
              value={fd.CorrectiveAction || ''}
              onChange={(_, v) => updateField('CorrectiveAction', v)}
            />
          </div>
          <div>
            <TextField
              label="Preventive Action"
              multiline
              rows={3}
              disabled={ro}
              value={fd.PreventiveAction || ''}
              onChange={(_, v) => updateField('PreventiveAction', v)}
            />
          </div>
        </div>
        <div className={styles.fieldFull}>
          <TextField
            label="Action Plan"
            multiline
            rows={3}
            disabled={ro}
            value={fd.ActionPlan || ''}
            onChange={(_, v) => updateField('ActionPlan', v)}
          />
        </div>
        <div className={styles.fieldFull}>
          <TextField
            label="Evidence"
            multiline
            rows={2}
            disabled={ro}
            value={fd.Evidence || ''}
            onChange={(_, v) => updateField('Evidence', v)}
          />
        </div>
      </div>

      {/* ── Section 5: Status & Dates ─────────────────────────────────── */}
      <div className={styles.section}>
        <h4>Status &amp; Dates</h4>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="Status"
              required
              disabled={ro}
              selectedKey={fd.Status || undefined}
              options={toDropdownOptions(Object.values(AuditStatus))}
              onChange={(_, opt) => updateField('Status', opt?.key)}
              errorMessage={errs.Status}
            />
          </div>
          <div>
            <Dropdown
              label="Priority"
              disabled={ro}
              selectedKey={fd.Priority || undefined}
              options={toDropdownOptions(Object.values(AuditPriority))}
              onChange={(_, opt) => updateField('Priority', opt?.key)}
            />
          </div>
          <div>
            <Dropdown
              label="Risk Rating"
              disabled={ro}
              selectedKey={fd.RiskRating || undefined}
              options={toDropdownOptions(Object.values(RiskRating))}
              onChange={(_, opt) => updateField('RiskRating', opt?.key)}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <DatePicker
              label="Audit Date"
              isRequired
              disabled={ro}
              value={fd.AuditDate ? new Date(fd.AuditDate) : undefined}
              onSelectDate={(date) => updateField('AuditDate', date ? date.toISOString() : '')}
              placeholder="Select audit date"
            />
            {errs.AuditDate && <span className={styles.errorText}>{errs.AuditDate}</span>}
          </div>
          <div>
            <DatePicker
              label="Due Date"
              disabled={ro}
              value={fd.DueDate ? new Date(fd.DueDate) : undefined}
              onSelectDate={(date) => updateField('DueDate', date ? date.toISOString() : '')}
              placeholder="Select due date"
            />
            {errs.DueDate && <span className={styles.errorText}>{errs.DueDate}</span>}
          </div>
          <div>
            <DatePicker
              label="Target Close Date"
              disabled={ro}
              value={fd.TargetCloseDate ? new Date(fd.TargetCloseDate) : undefined}
              onSelectDate={(date) => updateField('TargetCloseDate', date ? date.toISOString() : '')}
              placeholder="Select target close date"
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <DatePicker
              label="Completion Date"
              disabled={ro}
              value={fd.CompletionDate ? new Date(fd.CompletionDate) : undefined}
              onSelectDate={(date) => updateField('CompletionDate', date ? date.toISOString() : '')}
            />
          </div>
          <div>
            <TextField
              label="Extension Count"
              type="number"
              disabled={ro}
              value={fd.ExtensionCount !== undefined ? String(fd.ExtensionCount) : '0'}
              onChange={(_, v) => updateField('ExtensionCount', v ? parseInt(v, 10) : 0)}
            />
          </div>
        </div>
      </div>

      {/* ── Section 6: Verification & Follow-up ──────────────────────── */}
      <div className={styles.section}>
        <h4>Verification &amp; Follow-up</h4>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="Verification Result"
              disabled={ro}
              selectedKey={fd.VerificationResult || undefined}
              options={toDropdownOptions(Object.values(VerificationResult))}
              onChange={(_, opt) => updateField('VerificationResult', opt?.key)}
            />
          </div>
          <div>
            <DatePicker
              label="Verification Date"
              disabled={ro}
              value={fd.VerificationDate ? new Date(fd.VerificationDate) : undefined}
              onSelectDate={(date) => updateField('VerificationDate', date ? date.toISOString() : '')}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <Toggle
              label="Requires Follow-up"
              disabled={ro}
              checked={!!fd.RequiresFollowUp}
              onChange={(_, checked) => updateField('RequiresFollowUp', checked)}
            />
          </div>
        </div>
        <div className={styles.fieldFull}>
          <TextField
            label="Follow-up Notes"
            multiline
            rows={3}
            disabled={ro}
            value={fd.FollowUpNotes || ''}
            onChange={(_, v) => updateField('FollowUpNotes', v)}
          />
        </div>
        <div className={styles.fieldFull}>
          <TextField
            label="Remarks"
            multiline
            rows={2}
            disabled={ro}
            value={fd.Remarks || ''}
            onChange={(_, v) => updateField('Remarks', v)}
          />
        </div>
      </div>

      {/* ── Read-only calculated fields (edit mode only) ─────────────── */}
      {isEditMode && (
        <div className={styles.section}>
          <h4>Calculated Fields (Read-only)</h4>
          <div className={styles.fieldRow}>
            <div>
              <Label>Days Open</Label>
              <span>{fd.DaysOpen ?? 'N/A'}</span>
            </div>
            <div>
              <Label>Is Overdue</Label>
              <span>{fd.IsOverdue ? 'Yes' : 'No'}</span>
            </div>
          </div>
        </div>
      )}

      {/* ── Action buttons ────────────────────────────────────────────── */}
      {!ro && (
        <div className={styles.buttonBar}>
          <DefaultButton text="Cancel" onClick={onCancel} disabled={state.isSaving} />
          <PrimaryButton
            text={state.isSaving ? 'Saving...' : (isEditMode ? 'Update Item' : 'Create Item')}
            onClick={handleSubmit}
            disabled={state.isSaving}
          />
        </div>
      )}
      {ro && (
        <div className={styles.buttonBar}>
          <DefaultButton text="Back" onClick={onCancel} />
        </div>
      )}
    </div>
  );
};

export default AuditItemForm;
