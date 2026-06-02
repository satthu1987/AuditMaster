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
  FindingType,
  VerificationResult,
  InternalExternal,
  ActionStatus,
  QLVerification,
  RegionChoices
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
 * Matches the exact 34-field structure from the SharePoint list.
 *
 * Sections:
 *  1. Basic Information (Title, FindingType, Region, Internal/External, PIONumber)
 *  2. Classification (ISO clause, Finding ISO Chapter, Service, Category)
 *  3. People (PIC, Quality Manager, Auditor, Verified by)
 *  4. Finding Details (Finding Description, Root Cause, Quick fix, Action Taken)
 *  5. Dates & Status (Status, Action Status, dates, Required RCA)
 *  6. Verification & Links (Verification Result, Q&L verification, Evidence, Confluence Page)
 *  7. Calculated / Dependent Fields (read-only)
 */
const AuditItemForm: React.FC<IAuditItemFormProps> = (props) => {
  const { spService, roleService, currentUser, editItem, onSaved, onCancel } = props;

  const isEditMode = !!editItem;

  const [state, setState] = React.useState<IFormState>({
    formData: editItem ? { ...editItem } : {
      Status: AuditStatus.Open,
      VerificationResult: VerificationResult.No,
      ActionStatus: ActionStatus.Onprogress,
      QLVerification: QLVerification.No,
      RequiredRCA: false
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
        const errorMessage = err instanceof Error ? err.message : 'Failed to save audit item.';
        setState(prev => ({
          ...prev,
          isLoading: false,
          errorMessage: `Failed to load form data: ${errorMessage}`
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

  const segmentServiceOptions: IDropdownOption[] = state.segmentServices.map(s => ({
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
    if (!fd.Status) {
      errs.Status = 'Status is required.';
    }
    if (!fd.FindingDescription || fd.FindingDescription.trim() === '') {
      errs.FindingDescription = 'Finding Description is required.';
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
        await spService.createAuditItem(state.formData);
        setState(prev => ({
          ...prev,
          isSaving: false,
          successMessage: 'Audit item created successfully!'
        }));
      }

      setTimeout(() => onSaved(), 1200);
    } catch (err) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to save audit item.';
      setState(prev => ({
        ...prev,
        isSaving: false,
        errorMessage
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
        {isEditMode ? `Edit Audit Item – ${fd.Title || ''}` : 'Create New Audit Item'}
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
        <div className={styles.fieldFull}>
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
        </div>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="PIONumber"
              disabled={ro}
              value={fd.PIONumber || ''}
              onChange={(_, v) => updateField('PIONumber', v)}
              placeholder="e.g. PIO-2026-001"
            />
          </div>
          <div>
            <Dropdown
              label="Finding Type"
              disabled={ro}
              selectedKey={fd.FindingType || undefined}
              options={toDropdownOptions(Object.values(FindingType))}
              onChange={(_, opt) => updateField('FindingType', opt?.key)}
            />
          </div>
          <div>
            <Dropdown
              label="Region"
              disabled={ro}
              selectedKey={fd.Region || undefined}
              options={toDropdownOptions(RegionChoices)}
              onChange={(_, opt) => updateField('Region', opt?.key)}
            />
          </div>
          <div>
            <Dropdown
              label="Internal / External"
              disabled={ro}
              selectedKey={fd.InternalExternal || undefined}
              options={toDropdownOptions(Object.values(InternalExternal))}
              onChange={(_, opt) => updateField('InternalExternal', opt?.key)}
            />
          </div>
        </div>
        <div className={styles.fieldFull}>
          <TextField
            label="Finding Description"
            required
            multiline
            rows={4}
            disabled={ro}
            value={fd.FindingDescription || ''}
            onChange={(_, v) => updateField('FindingDescription', v)}
            errorMessage={errs.FindingDescription}
          />
        </div>
      </div>

      {/* ── Section 4: Finding Details ────────────────────────────────── */}
      <div className={styles.section}>
        <h4>Finding Details</h4>
        <div className={styles.fieldFull}>
          <TextField
            label="RC Description (Root Cause)"
            multiline
            rows={3}
            disabled={ro}
            value={fd.RCDescription || ''}
            onChange={(_, v) => updateField('RCDescription', v)}
          />
        </div>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Quick Fix"
              multiline
              rows={3}
              disabled={ro}
              value={fd.QuickFix || ''}
              onChange={(_, v) => updateField('QuickFix', v)}
            />
          </div>
          <div>
            <TextField
              label="Action Taken"
              multiline
              rows={3}
              disabled={ro}
              value={fd.ActionTaken || ''}
              onChange={(_, v) => updateField('ActionTaken', v)}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <Toggle
              label="Required RCA"
              disabled={ro}
              checked={!!fd.RequiredRCA}
              onChange={(_, checked) => updateField('RequiredRCA', checked)}
            />
          </div>
        </div>
      </div>

      {/* ── Section 2: Classification (Lookups) ──────────────────────── */}
      <div className={styles.section}>
        <h4>Classification</h4>
        <div className={styles.fieldRow}>
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
          <div>
            <Dropdown
              label="Finding ISO Chapter"
              disabled={ro}
              selectedKey={fd.FindingISOChapterId || undefined}
              options={isoClauseOptions}
              onChange={(_, opt) => updateField('FindingISOChapterId', opt?.key)}
              placeholder="Select finding ISO chapter"
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="Service"
              disabled={ro}
              selectedKey={fd.ServiceId || undefined}
              options={segmentServiceOptions}
              onChange={(_, opt) => updateField('ServiceId', opt?.key)}
              placeholder="Select service (auto-fills Segment &amp; Division)"
            />
          </div>
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
        </div>
        {/* Dependent lookup values (read-only) */}
        {isEditMode && (
          <div className={styles.fieldRow}>
            <div>
              <Label>Segment (auto)</Label>
              <span>{fd.Segment || '–'}</span>
            </div>
            <div>
              <Label>Article (auto)</Label>
              <span>{fd.Article || '–'}</span>
            </div>
            <div>
              <Label>Division (auto)</Label>
              <span>{fd.Service_Division || '–'}</span>
            </div>
          </div>
        )}
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
              description="SharePoint User ID of the Person In Charge"
              type="number"
            />
          </div>
          <div>
            <TextField
              label="Quality Manager (User ID)"
              disabled={ro}
              value={fd.QualityManagerId ? String(fd.QualityManagerId) : ''}
              onChange={(_, v) => updateField('QualityManagerId', v ? parseInt(v, 10) : undefined)}
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
              label="Verified by (User ID)"
              disabled={ro}
              value={fd.VerifiedById ? String(fd.VerifiedById) : ''}
              onChange={(_, v) => updateField('VerifiedById', v ? parseInt(v, 10) : undefined)}
              type="number"
            />
          </div>
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
              label="Action Status"
              disabled={ro}
              selectedKey={fd.ActionStatus || undefined}
              options={toDropdownOptions(Object.values(ActionStatus))}
              onChange={(_, opt) => updateField('ActionStatus', opt?.key)}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <DatePicker
              label="Audit Date"
              disabled={ro}
              value={fd.AuditDate ? new Date(fd.AuditDate) : undefined}
              onSelectDate={(date) => updateField('AuditDate', date ? date.toISOString() : '')}
              placeholder="Select audit date"
            />
          </div>
          <div>
            <DatePicker
              label="Due Date"
              disabled={ro}
              value={fd.DueDate ? new Date(fd.DueDate) : undefined}
              onSelectDate={(date) => updateField('DueDate', date ? date.toISOString() : '')}
              placeholder="Select due date"
            />
          </div>
          <div>
            <DatePicker
              label="Closed Date"
              disabled={ro}
              value={fd.ClosedDate ? new Date(fd.ClosedDate) : undefined}
              onSelectDate={(date) => updateField('ClosedDate', date ? date.toISOString() : '')}
            />
          </div>
        </div>
      </div>

      {/* ── Section 6: Verification & Links ───────────────────────────── */}
      <div className={styles.section}>
        <h4>Verification &amp; Links</h4>
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
            <Dropdown
              label="Q&amp;L Verification"
              disabled={ro}
              selectedKey={fd.QLVerification || undefined}
              options={toDropdownOptions(Object.values(QLVerification))}
              onChange={(_, opt) => updateField('QLVerification', opt?.key)}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
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
            <TextField
              label="Evidence (link) – URL"
              disabled={ro}
              value={fd.EvidenceLink?.Url || ''}
              onChange={(_, v) => updateField('EvidenceLink', { Url: v || '', Description: v || '' })}
              placeholder="https://..."
            />
          </div>
          <div>
            <TextField
              label="Confluence Page – URL"
              disabled={ro}
              value={fd.ConfluencePage?.Url || ''}
              onChange={(_, v) => updateField('ConfluencePage', { Url: v || '', Description: v || '' })}
              placeholder="https://..."
            />
          </div>
        </div>
      </div>

      {/* ── Section 7: Calculated & Dependent Fields (read-only, edit mode) ─ */}
      {isEditMode && (
        <div className={styles.section}>
          <h4>Calculated Fields (Read-only)</h4>
          <div className={styles.fieldRow}>
            <div>
              <Label>Year</Label>
              <span>{fd.Year || 'N/A'}</span>
            </div>
            <div>
              <Label>Finding Number</Label>
              <span>{fd.FindingNumber || 'N/A'}</span>
            </div>
            <div>
              <Label>Verification Date (Calculated)</Label>
              <span>{fd.VerificationDateCalculated || 'N/A'}</span>
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
