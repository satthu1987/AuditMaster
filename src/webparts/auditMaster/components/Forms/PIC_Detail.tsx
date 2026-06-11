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
  IValidationErrors,
  AuditStatus,
  FindingType,
  VerificationResult,
  QLVerification
} from '../../models';
import { RichTextEditor } from '../Common';
import { IAuditItemFormProps } from './AuditItemForm';

/**
 * ════════════════════════════════════════════════════════════════════════════
 *  PIC_Detail – PIC (Person In Charge) specific audit detail form
 * ════════════════════════════════════════════════════════════════════════════
 *
 *  This is a role-tailored variant of `AuditItemForm`, used by `PICView` for the
 *  PIC role. It shares the exact same props interface (`IAuditItemFormProps`) so
 *  it is a drop-in replacement, and reuses the same shared components
 *  (RolePeoplePicker, RichTextEditor) and SharePoint services.
 *
 *  EDITABILITY FOR A PIC
 *  ─────────────────────
 *    ✔ EDITABLE   – Finding Details  (description, root-cause, quick fix,
 *                                      action taken, required-RCA toggle)
 *    ✔ EDITABLE   – Status field      (so the PIC can move the item forward,
 *                                      e.g. → Pending Verification)
 *    ✔ EDITABLE   – Verification & Links (the PIC prepares evidence / links)
 *
 *    RO READ-ONLY – Basic Information   (set when the finding was raised)
 *    RO READ-ONLY – Classification      (ISO clause, service, category …)
 *    RO READ-ONLY – People              (PIC / QM / Auditor / Verifier)
 *    RO READ-ONLY – Action Status + all Dates (admin/verifier controlled)
 *    RO READ-ONLY – Calculated fields
 *
 *  The descriptive sections are still rendered (disabled) so the PIC has full
 *  context while working their finding. If the PIC has no edit rights on the
 *  item at all (RoleService.canUpdate === false, e.g. the item is not assigned
 *  to them) the whole form falls back to read-only via the `ro` flag.
 * ════════════════════════════════════════════════════════════════════════════
 */

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
  // Cascading ISO Clause derived values (read-only context for the PIC)
  isoChapter: string;
  isoLevel: string;
  isoArticle: string;
}

const PIC_Detail: React.FC<IAuditItemFormProps> = (props) => {
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
    errors: {},
    isSaving: false,
    isLoading: true,
    successMessage: '',
    errorMessage: '',
    isReadOnly: false,
    isoChapter: '',
    isoLevel: '',
    isoArticle: ''
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

        // Permission check – if the PIC can't update this specific item the
        // whole form is shown read-only.
        let readOnly = false;
        if (isEditMode && editItem) {
          readOnly = !roleService.canUpdate(currentUser, editItem);
        } else {
          readOnly = !roleService.canCreate(currentUser);
        }

        // Pre-populate read-only derived ISO Clause display values.
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
          isLoading: false,
          isReadOnly: readOnly,
          isoChapter,
          isoLevel,
          isoArticle
        }));
      } catch (err) {
        const errorMessage = err instanceof Error ? err.message : 'Failed to load form data.';
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

    if (!fd.DueDate) {
      errs.DueDate = 'Due Date is required.';
    }

    const actionTakenValue = (fd.ActionTaken || '').replace(/<[^>]*>/g, '').trim();
    if (!actionTakenValue) {
      errs.ActionTaken = 'Action Taken is required.';
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
        const payload = {
          ...state.formData,
          Status: AuditStatus.InProgress
        };
        await spService.updateAuditItem(editItem.Id, payload);
        setState(prev => ({
          ...prev,
          formData: { ...prev.formData, Status: AuditStatus.InProgress },
          isSaving: false,
          successMessage: 'Audit item updated successfully!'
        }));
      } 

      setTimeout(() => onSaved(), 1200);
    } catch (err) {
        const errorMessage = err instanceof Error ? err.message : 'Failed to save audit item.';
      setState(prev => ({
        ...prev,
        isSaving: false,
        errorMessage: `Save failed: ${errorMessage}`
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
  const selectedISOClause = state.isoClauses.find(c => c.Id === (fd.ISOClauseId || (fd as Partial<IAuditMasterItem> & { ISO_x0020_clauseId?: number }).ISO_x0020_clauseId));

  // ══════════════════════════════════════════════════════════════════════════
  //  PIC EDITABILITY MAP
  //  `true`  → the field/section is disabled (read-only)
  //  `ro`    → editable only when the PIC has update rights on this item
  // ══════════════════════════════════════════════════════════════════════════
  const sectionReadOnly = {
    basicInfo: true,
    classification: true,
    people: true,
    findingDetails: true,
    status: true,
    dates: false,
    verification: true
  };

  return (
    <div className={styles.formContainer}>
      <div className={styles.formTitle}>
        {`Audit Item – ${fd.Title || ''}`}
      </div>
      <div className={styles.formSubtitle}>
        PIC view – update Action Taken and Due Date.
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
      {ro ? (
        <div className={styles.readOnlyBanner}>
          You do not have permission to edit this item. Displaying in read-only mode.
        </div>
      ) : (
        <MessageBar messageBarType={MessageBarType.info} isMultiline={true}>
          As the PIC you must fill <strong>Action Taken</strong> and <strong>Due Date</strong> to update this item.
        </MessageBar>
      )}

      {/* ── Section 1: Basic Information ──────────────────────────────── */}
      <div className={styles.section}>
        <h4>Basic Information</h4>
        <div className={styles.fieldFull}>
          <div>
            <TextField
              label="Title"
              required
              disabled={true}
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
              disabled
              value={fd.PIONumber || ''}
              onChange={(_, v) => updateField('PIONumber', v)}
              placeholder="e.g. PIO-2026-001"
            />
          </div>
          <div>
            <TextField
              label="Finding Type"
              disabled
              readOnly
              value={fd.FindingType || ''}
            />
          </div>
          <div>
             <TextField
              label="Audit Type"
              disabled
              readOnly
              value={fd.AuditType || ''}
            />
          </div>
        </div>
        <div className={styles.fieldFull}>
          <RichTextEditor
            label="Finding Description"
            minHeight={160}
            disabled
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
            <TextField
              label="ISO Clause"
              disabled
              readOnly
              value={selectedISOClause?.ISOClause || fd?.ISOClause?.ISOClause || ''}
            />
          </div>
          <div>
            <TextField
              label="ISO Chapter"
              readOnly
              disabled
              value={state.isoChapter || selectedISOClause?.ISOlevel2 || fd?.ISOClause?.ISOlevel2 || ''}
            />
          </div>
          <div>
            <TextField
              label="ISO Level"
              readOnly
              disabled
              value={state.isoLevel || selectedISOClause?.ISOlevel1 || fd?.ISOClause?.ISOlevel1 || ''}
            />
          </div>
          <div>
            <TextField
              label="Article"
              readOnly
              disabled
              value={fd?.Article || state.isoArticle || selectedISOClause?.Article || fd?.ISOClause?.Article || ''}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Segment"
              readOnly
              disabled
              value={fd?.Segment}
            />
          </div>
          <div>
            <TextField
              label="Service"
              readOnly
              disabled
              value={fd?.Service}
            />
          </div>
          <div>
            <TextField
              label="Division"
              readOnly
              disabled
              value={fd?.Division}
            />
          </div>
        </div>
        
      </div>

      {/* ── Section 4: Finding Details ────────────────────────────────── */}
      <div className={styles.section}>
        <h4>Finding Details</h4>
        <div className={styles.fieldFull}>
          <RichTextEditor
            label="RC Description (Root Cause)"
            disabled
            value={fd.RCDescription || ''}
            onChange={(html:string) => updateField('RCDescription', html)}
          />
        </div>
        <div className={styles.fieldRow}>
          <div>
            <RichTextEditor
              label="Quick Fix"
              disabled
              value={fd.QuickFix || ''}
              onChange={(html:string) => updateField('QuickFix', html)}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <RichTextEditor
              label="Action Taken"
              required
              disabled={ro}
              value={fd.ActionTaken || ''}
              onChange={(html:string) => updateField('ActionTaken', html)}
              minHeight={140}
            />
            {errs.ActionTaken && <span style={{ color: '#a4262c', fontSize: 12 }}>{errs.ActionTaken}</span>}
          </div>
        </div>
      </div>


      {/* ── Section 3: People ─────────────────────────────────────────── */}
      <div className={styles.section}>
        <h4>People</h4>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="PIC"
              readOnly
              disabled
              value={fd?.PIC?.Title}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Auditor"
              readOnly
              disabled
              value={fd?.Auditor?.Title}
            />
          </div>
          <div>
            <TextField
              label="Verified by"
              readOnly
              disabled
              value={fd?.Verifier?.Title}
            />
          </div>
        </div>
      </div>

      {/* ── Section 5: Status & Dates ─────────────────────────────────── */}
      <div className={styles.section}>
        <h4>Status &amp; Dates</h4>
        <div className={styles.fieldRow}>
          <div>
            <TextField
              label="Status"
              readOnly
              disabled
              value={fd?.Status}
            />
          </div>
          <div>
            <TextField
              label="Audit Date"
              readOnly
              disabled
              value={fd?.AuditDate}
            />
          </div>
          <div>
            <DatePicker
              label="Due Date"
              isRequired={true}
              disabled={ro}
              value={fd?.DueDate ? new Date(fd.DueDate) : undefined}
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
              disabled
              selectedKey={fd?.VerificationResult || undefined}
              options={toDropdownOptions(Object.values(VerificationResult))}
              onChange={(_, opt) => updateField('VerificationResult', opt?.key)}
            />
          </div>
          <div>
            <Dropdown
              label="Q&L Verification"
              disabled
              selectedKey={fd?.QLVerification || undefined}
              options={toDropdownOptions(Object.values(VerificationResult))}
              onChange={(_, opt) => updateField('QLVerification', opt?.key)}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <DatePicker
              label="Verification Date"
              disabled
              value={fd?.VerificationDate ? new Date(fd.VerificationDate) : undefined}
              onSelectDate={(date) => updateField('VerificationDate', date ? date.toISOString() : '')}
            />
          </div>
        </div>
        <div className={styles.fieldRow}>
          <div>
            <RichTextEditor
              label="Evidence (link) – URL"
              disabled
              value={fd?.EvidenceLink?.Url || ''}
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
              <span>{fd?.Year || 'N/A'}</span>
            </div>
            <div>
              <Label>Verification Date (Calculated)</Label>
              <span>{fd?.VerificationDateCalculated || 'N/A'}</span>
            </div>
          </div>
        </div>
      )}

      {/* ── Action buttons ──────────────────────────────────────────────── */}
      {!ro && (
        <div className={styles.buttonBar}>
          <DefaultButton text="Cancel" onClick={onCancel} disabled={state.isSaving} />
          <PrimaryButton
            text={state.isSaving ? 'Saving...' : 'Update Item'}
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

export default PIC_Detail;
