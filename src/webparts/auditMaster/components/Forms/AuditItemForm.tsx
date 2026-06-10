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
  IUserRoleItem,
  AuditStatus,
  FindingType,
  AuditType,
  VerificationResult,
  QLVerification
} from '../../models';
import { SharePointService, RoleService } from '../../services';
import { RichTextEditor } from '../Common';

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
  userRoles: IUserRoleItem[];
  usersById: { [id: number]: { Id: number; Title: string; Email: string } };
  errors: IValidationErrors;
  isSaving: boolean;
  isLoading: boolean;
  successMessage: string;
  errorMessage: string;
  isReadOnly: boolean;
  // Cascading ISO Clause derived values (auto-populated from the selected clause)
  isoChapter: string;   // ← ISOlevel2
  isoLevel: string;     // ← ISOlevel1
  isoArticle: string;   // ← Article
  isLoadingISO: boolean;
}

/**
 * Full Audit Master item form – used for both Create and Edit.
 * Matches the exact 34-field structure from the SharePoint list.
 *
 * Sections:
 *  1. Basic Information (Title, FindingType, Region, Internal/External, PIONumber)
 *  2. Classification (ISO clause, Service, Category)
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
      QLVerification: QLVerification.No
    },
    categories: [],
    isoClauses: [],
    segmentServices: [],
    userRoles: [],
    usersById: {},
    errors: {},
    isSaving: false,
    isLoading: true,
    successMessage: '',
    errorMessage: '',
    isReadOnly: false,
    isoChapter: '',
    isoLevel: '',
    isoArticle: '',
    isLoadingISO: false
  });

  // ── Load reference data on mount ──────────────────────────────────────────
  React.useEffect(() => {
    const loadData = async (): Promise<void> => {
      try {
        const [categories, isoClauses, segmentServices, userRoles] = await Promise.all([
          spService.getCategories(),
          spService.getISOClauses(),
          spService.getSegmentServices(),
          spService.getUserRoles()
        ]);

        const roleUserIds = Array.from(
          new Set(
            userRoles
              .filter(r => r.AccountId)
              .map(r => r.AccountId)
          )
        );

       

        // Pre-populate derived ISO Clause fields when editing an existing item
        let isoChapter = '';
        let isoLevel = '';
        let isoArticle = '';
        const existingClauseId = editItem?.ISOClauseId;
        if (existingClauseId) {
          const clause = isoClauses.filter(c => c.Id === existingClauseId)[0];
          if (clause) {
            isoChapter = clause.ISOlevel2 || '';
            isoLevel = clause.ISOlevel1 || '';
            isoArticle = clause.Article || '';
          }
        }

        setState(prev => ({
          ...prev,
          categories,
          isoClauses,
          segmentServices,
          userRoles,
          isLoading: false,
          isoChapter,
          isoLevel,
          isoArticle
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

 // ── Cascading ISO Clause selection ─────────────────────────────────────────
  // When the ISO Clause changes, fetch the full ISOClause list item and
  // auto-populate the read-only ISO Chapter (ISOlevel2), ISO Level (ISOlevel1)
  // and Article fields. Clearing the clause resets all three.
  const handleISOClauseChange = async (clauseId: number | undefined): Promise<void> => {
    // Update the lookup id first
    setState(prev => ({
      ...prev,
      formData: { ...prev.formData, ISO_x0020_clauseId: clauseId as number },
      errors: { ...prev.errors, ISO_x0020_clauseId: '' },
      successMessage: '',
      errorMessage: ''
    }));

    // Cleared → reset dependent fields. (Article/ISO Chapter/ISO Level are
    // derived display values, so they are reset in local state only.)
    if (!clauseId) {
      setState(prev => ({
        ...prev,
        isoChapter: '',
        isoLevel: '',
        isoArticle: '',
        isLoadingISO: false,
        formData: { ...prev.formData, ISO_x0020_clauseId: undefined as any }
      }));
      return;
    }

    // Start loading the dependent values
    setState(prev => ({ ...prev, isLoadingISO: true }));

    try {
      const clause = await spService.getISOClauseById(clauseId);
      setState(prev => ({
        ...prev,
        isLoadingISO: false,
        isoChapter: clause.ISOlevel2 || '',
        isoLevel: clause.ISOlevel1 || '',
        isoArticle: clause.Article || ''
      }));
    } catch (err) {
      // Fallback to the cached list if the single-item fetch fails
      const cached = state.isoClauses.filter(c => c.Id === clauseId)[0];
      const errorMessage = err instanceof Error ? err.message : 'Failed to load ISO Clause details.';
      setState(prev => ({
        ...prev,
        isLoadingISO: false,
        isoChapter: cached ? (cached.ISOlevel2 || '') : '',
        isoLevel: cached ? (cached.ISOlevel1 || '') : '',
        isoArticle: cached ? (cached.Article || '') : '',
        errorMessage: cached ? '' : `Failed to load ISO Clause details: ${errorMessage}`
      }));
    }
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

  const isoChapterOptions: IDropdownOption[] = Array.from(
    new Map(
      state.isoClauses.map(c => [c.ISOlevel1 || c.ISOClause, c] as const)
    ).entries()
  ).map(([, c]) => ({
    key: c.Id,
    text: c.ISOlevel1 || c.ISOClause
  }));

  const isoArticleOptions: IDropdownOption[] = Array.from(
    new Map(
      state.isoClauses.map(c => [c.Article || c.ISOClause, c] as const)
    ).entries()
  ).map(([, c]) => ({
    key: c.Id,
    text: c.Article
  }));

  const segmentOptions: IDropdownOption[] = Array.from(
    new Set(state.segmentServices.map(s => s.Title).filter(Boolean))
  ).map(segment => ({
    key: segment,
    text: segment
  }));

  const serviceOptions: IDropdownOption[] = state.segmentServices
    .filter(s => !state.formData.Segment || s.Title === state.formData.Segment)
    .map(s => ({
      key: s.Service || s.Title,
      text: s.Service || s.Title
    }));

    const divisionOptions: IDropdownOption[] = Array.from(
    new Set(
      state.segmentServices
        .filter(s => !state.formData.Segment || s.Title === state.formData.Segment)
        .map(s => (s.Division || s.Title || '').trim())
        .filter(v => v.length > 0)
    )
  ).map(division => ({
    key: division,
    text: division
  }));


  const getUserRoleOptions = (allowedRoles: string[]): IDropdownOption[] => Array.from(
    new Map(
      state.userRoles
        .filter(r => r.AccountId && allowedRoles.indexOf(String(r.Role)) !== -1)
        .map(r => [r.AccountId, r] as const)
    ).values()
  ).map(r => {
    const user = state.usersById[r.AccountId];
    const accountLabel = r.AccountStringId && r.AccountStringId.indexOf('|') !== -1
      ? r.AccountStringId.split('|').pop()
      : r.AccountStringId;

    const text = user
      ? `${user.Title}${user.Email ? ` (${user.Email})` : ''}`
      : (accountLabel || `User ${r.AccountId}`);

    return {
      key: r.AccountId,
      text
    };
  });

  const picOptions: IDropdownOption[] = getUserRoleOptions(['PIC']);
  const qualityManagerOptions: IDropdownOption[] = getUserRoleOptions(['Manager']);
  const verifierOptions: IDropdownOption[] = getUserRoleOptions(['Verifier']);
  const auditorOptions: IDropdownOption[] = getUserRoleOptions(['Auditor']);
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
    if (!fd.FindingType) {
      errs.FindingType = 'Finding Type is required.';
    }
    if (!fd.ISOClauseId) {
      errs.ISOClauseId = 'ISO Clause is required.';
    }
    if (!fd.Segment) {
      errs.Segment = 'Segment is required.';
    }
    if (!fd.Service) {
      errs.Service = 'Service is required.';
    }
    if (!fd.CategoryId) {
      errs.CategoryId = 'Category is required.';
    }
    if (!fd.PICId) {
      errs.PICId = 'PIC is required.';
    }
    if (!fd.AuditorId) {
      errs.AuditorId = 'Auditor is required.';
    }
    if (!fd.AuditDate) {
      errs.AuditDate = 'Audit Date is required.';
    }
    if (!fd.DueDate) {
      errs.DueDate = 'Due Date is required.';
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
              required
              disabled={ro}
              selectedKey={fd.FindingType || undefined}
              options={toDropdownOptions(Object.values(FindingType))}
              onChange={(_, opt) => updateField('FindingType', opt?.key)}
              errorMessage={errs.FindingType}
            />
          </div>
          <div>
            <Dropdown
              label="Audit Type"
              disabled={ro}
              selectedKey={fd.AuditType || undefined}
              options={toDropdownOptions(Object.values(AuditType))}
              onChange={(_, opt) => updateField('AuditType', opt?.key)}
              placeholder="Select audit type"
            />
          </div>
        </div>
        <div className={styles.fieldFull}>
          <RichTextEditor
            label="Finding Description"
            required
            minHeight={160}
            disabled={ro}
            value={fd.FindingDescription}
            onChange={(html: string) => updateField('FindingDescription', html)}
          />
        </div>
      </div>

      {/* ── Section 2: Classification (Lookups) ──────────────────────── */}
      <div className={styles.section}>
        <h4>Classification</h4>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="ISO Clause"
              required
              disabled={ro}
              selectedKey={fd.ISOClauseId || undefined}
              options={isoClauseOptions}
              onChange={(_, opt) => handleISOClauseChange(opt?.key as number | undefined)}
              placeholder="Select ISO clause"
              errorMessage={errs.ISOClauseId}
            />
          </div>
          <div>
            <TextField
              label="ISO Chapter"
              readOnly
              disabled
              value={state.isoChapter}
              placeholder="Auto-filled from ISO Clause"
              description="Auto-populated from the selected ISO Clause (ISOlevel2)"
            />
          </div>
          <div>
            <TextField
              label="ISO Level"
              readOnly
              disabled
              value={state.isoLevel}
              placeholder="Auto-filled from ISO Clause"
              description="Auto-populated from the selected ISO Clause (ISOlevel1)"
            />
          </div>
          <div>
            <TextField
              label="Article"
              readOnly
              disabled
              value={state.isoArticle}
              placeholder="Auto-filled from ISO Clause"
              description="Auto-populated from the selected ISO Clause"
            />
            {state.isLoadingISO && (
              <div style={{ marginTop: 4 }}>
                <Spinner size={SpinnerSize.xSmall} label="Loading ISO details..." labelPosition="right" />
              </div>
            )}
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="Segment"
              required
              disabled={ro}
              selectedKey={fd.Segment || undefined}
              options={segmentOptions}
              onChange={(_, opt) => {
                updateField('Segment', opt?.key);
                updateField('Service', undefined);
              }}
              placeholder="Select segment"
              errorMessage={errs.Segment}
            />
          </div>
          <div>
            <Dropdown
              label="Service"
              required
              disabled={ro || !fd.Segment}
              selectedKey={fd.Service || undefined}
              options={serviceOptions}
              onChange={(_, opt) => {
                const selected = state.segmentServices.find(s => (s.Service || s.Title) === String(opt?.key));
                updateField('Service', opt?.key);
                updateField('Segment', selected ? selected.Title : fd.Segment);
              }}
              placeholder={fd.Segment ? 'Select service' : 'Select segment first'}
              errorMessage={errs.Service}
            />
          </div>
          <div>
            <Dropdown
              label="Division"
              required
              disabled={ro || !fd.Segment}
              selectedKey={fd.Division || undefined}
              options={divisionOptions}
              onChange={(_, opt) => {
                const selected = state.segmentServices.find(s => (s.Division || s.Title) === String(opt?.key));
                updateField('Division', opt?.key);
                updateField('Segment', selected ? selected.Title : fd.Segment);
              }}
              placeholder={fd.Segment ? 'Select division' : 'Select segment first'}
              errorMessage={errs.Division}
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
          </div>
        )}
      </div>

      {/* ── Section 4: Finding Details ────────────────────────────────── */}
      <div className={styles.section}>
        <h4>Finding Details</h4>
        <div className={styles.fieldFull}>
          <RichTextEditor
            label="RC Description (Root Cause)"
            disabled={ro}
            value={fd.RCDescription || ''}
            onChange={(html:string) => updateField('RCDescription', html)}
          />
        </div>
        <div className={styles.fieldRow}>
          <div>
            <RichTextEditor
              label="Quick Fix"
              disabled={ro}
              value={fd.QuickFix || ''}
              onChange={(html:string) => updateField('QuickFix', html)}
            />
          </div>
        </div>
      </div>


      {/* ── Section 3: People ─────────────────────────────────────────── */}
      <div className={styles.section}>
        <h4>People</h4>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="PIC"
              required
              disabled={ro}
              selectedKey={fd.PICId || undefined}
              options={picOptions}
              onChange={(_, opt) => updateField('PICId', opt?.key)}
              placeholder="Select PIC user or Admin"
              errorMessage={errs.PICId}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <Dropdown
              label="Auditor"
              required
              disabled={ro}
              selectedKey={fd.AuditorId || undefined}
              options={auditorOptions}
              onChange={(_, opt) => updateField('AuditorId', opt?.key)}
              placeholder="Select Auditor or Admin"
              errorMessage={errs.AuditorId}
            />
          </div>
          <div>
            <Dropdown
              label="Verified by"
              disabled={ro}
              selectedKey={fd.VerifierId || undefined}
              options={verifierOptions}
              onChange={(_, opt) => updateField('VerifierId', opt?.key)}
              placeholder="Select Verifier or Admin"
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
        </div>
        <div className={styles.fieldRow}>
          <div>
            <DatePicker
              label="Audit Date"
              isRequired={true}
              disabled={ro}
              value={fd.AuditDate ? new Date(fd.AuditDate) : undefined}
              onSelectDate={(date) => updateField('AuditDate', date ? date.toISOString() : '')}
              placeholder="Select audit date"
              textField={{ errorMessage: errs.AuditDate }}
            />
          </div>
          <div>
            <DatePicker
              label="Due Date"
              isRequired={true}
              disabled={ro}
              value={fd.DueDate ? new Date(fd.DueDate) : undefined}
              onSelectDate={(date) => updateField('DueDate', date ? date.toISOString() : '')}
              placeholder="Select due date"
              textField={{ errorMessage: errs.DueDate }}
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
              label="Q&L Verification"
              disabled={ro}
              selectedKey={fd.QLVerification || undefined}
              options={toDropdownOptions(Object.values(VerificationResult))}
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
            <RichTextEditor
              label="Evidence (link) – URL"
              disabled={ro}
              value={fd.EvidenceLink?.Url || ''}
              onChange={(html:string) => updateField('EvidenceLink', { Url: html || '', Description: html || '' })}
              placeholder="https://..."
              minHeight={168}
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
