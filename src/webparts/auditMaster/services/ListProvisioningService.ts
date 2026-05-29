import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  LIST_NAMES,
  DivisionChoices,
  DetailCategoryChoices,
  AuditStatus,
  AuditPriority,
  AuditType,
  AuditSource,
  VerificationResult,
  RiskRating,
  UserRole
} from '../models';

/**
 * Service responsible for ensuring all required SharePoint lists exist.
 * Creates lists and their fields if they do not already exist on the site.
 */
export class ListProvisioningService {
  private _context: WebPartContext;
  private _siteUrl: string;

  constructor(context: WebPartContext, siteUrl: string) {
    this._context = context;
    this._siteUrl = siteUrl;
  }

  // ─── Public API ──────────────────────────────────────────────────────────────

  /**
   * Provision all five lists. Safe to call multiple times – skips lists that
   * already exist.
   */
  public async provisionAllLists(): Promise<void> {
    console.log('[ListProvisioning] Starting list provisioning...');

    await this._provisionCategoryList();
    await this._provisionISOClauseList();
    await this._provisionSegmentServiceList();
    await this._provisionUserRoleList();
    await this._provisionAuditMasterList();

    console.log('[ListProvisioning] All lists provisioned successfully.');
  }

  // ─── Category List ─────────────────────────────────────────────────────────

  private async _provisionCategoryList(): Promise<void> {
    const listTitle = LIST_NAMES.CATEGORY;
    if (await this._listExists(listTitle)) {
      console.log(`[ListProvisioning] List "${listTitle}" already exists.`);
      return;
    }

    await this._createList(listTitle, 'Audit categories', 100);

    await this._addTextField(listTitle, 'Category', true);
    await this._addNoteField(listTitle, 'Notes', false);
    await this._addTextField(listTitle, 'CategoryGrouping1', false);

    console.log(`[ListProvisioning] List "${listTitle}" created.`);
  }

  // ─── ISO Clause List ───────────────────────────────────────────────────────

  private async _provisionISOClauseList(): Promise<void> {
    const listTitle = LIST_NAMES.ISO_CLAUSE;
    if (await this._listExists(listTitle)) {
      console.log(`[ListProvisioning] List "${listTitle}" already exists.`);
      return;
    }

    await this._createList(listTitle, 'ISO Clause reference data', 100);

    await this._addTextField(listTitle, 'ISOClause', true);
    await this._addTextField(listTitle, 'ISOlevel2', false);
    await this._addTextField(listTitle, 'ISOlevel1', false);
    await this._addTextField(listTitle, 'Article', false);
    await this._addNumberField(listTitle, 'Order', false);

    console.log(`[ListProvisioning] List "${listTitle}" created.`);
  }

  // ─── Segment_Service List ──────────────────────────────────────────────────

  private async _provisionSegmentServiceList(): Promise<void> {
    const listTitle = LIST_NAMES.SEGMENT_SERVICE;
    if (await this._listExists(listTitle)) {
      console.log(`[ListProvisioning] List "${listTitle}" already exists.`);
      return;
    }

    await this._createList(listTitle, 'Segment and service reference data', 100);

    await this._addTextField(listTitle, 'Service', false);
    await this._addTextField(listTitle, 'Service_Short', false);
    await this._addChoiceField(listTitle, 'Division', DivisionChoices, false);
    await this._addNumberField(listTitle, 'Order', false);
    await this._addChoiceField(listTitle, 'DetailCategory', DetailCategoryChoices, false);

    console.log(`[ListProvisioning] List "${listTitle}" created.`);
  }

  // ─── UserRole List ─────────────────────────────────────────────────────────

  private async _provisionUserRoleList(): Promise<void> {
    const listTitle = LIST_NAMES.USER_ROLE;
    if (await this._listExists(listTitle)) {
      console.log(`[ListProvisioning] List "${listTitle}" already exists.`);
      return;
    }

    await this._createList(listTitle, 'User role assignments', 100);

    await this._addPersonField(listTitle, 'Account');
    await this._addChoiceField(
      listTitle,
      'Role',
      [UserRole.Admin, UserRole.PIC, UserRole.Verifier],
      true
    );

    console.log(`[ListProvisioning] List "${listTitle}" created.`);
  }

  // ─── Audit Master List ─────────────────────────────────────────────────────

  private async _provisionAuditMasterList(): Promise<void> {
    const listTitle = LIST_NAMES.AUDIT_MASTER;
    if (await this._listExists(listTitle)) {
      console.log(`[ListProvisioning] List "${listTitle}" already exists.`);
      return;
    }

    await this._createList(listTitle, 'Main audit findings list', 100);

    // Text fields
    await this._addTextField(listTitle, 'AuditId', true);
    await this._addNoteField(listTitle, 'Description', false);
    await this._addTextField(listTitle, 'AuditYear', false);
    await this._addNoteField(listTitle, 'Finding', false);
    await this._addNoteField(listTitle, 'RootCause', false);
    await this._addNoteField(listTitle, 'CorrectiveAction', false);
    await this._addNoteField(listTitle, 'PreventiveAction', false);
    await this._addNoteField(listTitle, 'Evidence', false);
    await this._addNoteField(listTitle, 'Remarks', false);
    await this._addNoteField(listTitle, 'ActionPlan', false);
    await this._addNoteField(listTitle, 'FollowUpNotes', false);
    await this._addTextField(listTitle, 'Department', false);
    await this._addTextField(listTitle, 'Location', false);

    // Choice fields
    await this._addChoiceField(
      listTitle, 'Status',
      Object.values(AuditStatus),
      true, AuditStatus.Open
    );
    await this._addChoiceField(
      listTitle, 'Priority',
      Object.values(AuditPriority),
      false, AuditPriority.Medium
    );
    await this._addChoiceField(
      listTitle, 'AuditType',
      Object.values(AuditType),
      false
    );
    await this._addChoiceField(
      listTitle, 'Source',
      Object.values(AuditSource),
      false
    );
    await this._addChoiceField(
      listTitle, 'VerificationResult',
      Object.values(VerificationResult),
      false, VerificationResult.NotVerified
    );
    await this._addChoiceField(
      listTitle, 'RiskRating',
      Object.values(RiskRating),
      false, RiskRating.Medium
    );

    // Lookup fields
    const categoryListId = await this._getListId(LIST_NAMES.CATEGORY);
    if (categoryListId) {
      await this._addLookupField(listTitle, 'Category', categoryListId, 'Category');
    }

    const isoClauseListId = await this._getListId(LIST_NAMES.ISO_CLAUSE);
    if (isoClauseListId) {
      await this._addLookupField(listTitle, 'ISOClause', isoClauseListId, 'ISOClause');
    }

    const segmentListId = await this._getListId(LIST_NAMES.SEGMENT_SERVICE);
    if (segmentListId) {
      await this._addLookupField(listTitle, 'SegmentService', segmentListId, 'Title');
    }

    // Person or Group fields
    await this._addPersonField(listTitle, 'PIC');
    await this._addPersonField(listTitle, 'Verifier');
    await this._addPersonField(listTitle, 'Auditor');
    await this._addPersonField(listTitle, 'Auditee');
    await this._addPersonField(listTitle, 'CreatedByUser');

    // Date fields
    await this._addDateField(listTitle, 'AuditDate', false);
    await this._addDateField(listTitle, 'DueDate', false);
    await this._addDateField(listTitle, 'CompletionDate', false);
    await this._addDateField(listTitle, 'VerificationDate', false);
    await this._addDateField(listTitle, 'TargetCloseDate', false);

    // Number fields
    await this._addNumberField(listTitle, 'ExtensionCount', false);

    // Calculated fields
    await this._addCalculatedField(
      listTitle,
      'DaysOpen',
      '=IF([Status]="Closed",DATEDIF([AuditDate],[CompletionDate],"d"),DATEDIF([AuditDate],TODAY(),"d"))',
      'Number'
    );
    await this._addCalculatedField(
      listTitle,
      'IsOverdue',
      '=IF(AND([Status]<>"Closed",[DueDate]<TODAY()),"Yes","No")',
      'Text'
    );

    // Boolean fields
    await this._addBooleanField(listTitle, 'RequiresFollowUp', false);

    console.log(`[ListProvisioning] List "${listTitle}" created.`);
  }

  // ─── Helper Methods ────────────────────────────────────────────────────────

  private async _listExists(listTitle: string): Promise<boolean> {
    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        url, SPHttpClient.configurations.v1
      );
      return response.ok;
    } catch {
      return false;
    }
  }

  private async _getListId(listTitle: string): Promise<string | null> {
    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')?$select=Id`;
      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        url, SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const data = await response.json();
        return data.Id;
      }
      return null;
    } catch {
      return null;
    }
  }

  private async _createList(title: string, description: string, templateType: number): Promise<void> {
    const url = `${this._siteUrl}/_api/web/lists`;
    const body = JSON.stringify({
      '__metadata': { 'type': 'SP.List' },
      'Title': title,
      'Description': description,
      'BaseTemplate': templateType,
      'AllowContentTypes': true,
      'ContentTypesEnabled': false
    });

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': ''
      },
      body: body
    };

    const response = await this._context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to create list "${title}": ${errText}`);
    }
  }

  private async _addField(listTitle: string, fieldXml: string): Promise<void> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/createfieldasxml`;
    const body = JSON.stringify({
      'parameters': {
        '__metadata': { 'type': 'SP.XmlSchemaFieldCreationInformation' },
        'SchemaXml': fieldXml,
        'Options': 8 // AddFieldToDefaultView
      }
    });

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': ''
      },
      body: body
    };

    const response = await this._context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!response.ok) {
      const errText = await response.text();
      console.warn(`[ListProvisioning] Field creation warning for "${listTitle}": ${errText}`);
    }
  }

  private async _addTextField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="Text" DisplayName="${fieldName}" Name="${fieldName}" Required="${required ? 'TRUE' : 'FALSE'}" MaxLength="255" />`;
    await this._addField(listTitle, xml);
  }

  private async _addNoteField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="Note" DisplayName="${fieldName}" Name="${fieldName}" Required="${required ? 'TRUE' : 'FALSE'}" NumLines="6" RichText="TRUE" />`;
    await this._addField(listTitle, xml);
  }

  private async _addNumberField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="Number" DisplayName="${fieldName}" Name="${fieldName}" Required="${required ? 'TRUE' : 'FALSE'}" Decimals="0" Min="0" />`;
    await this._addField(listTitle, xml);
  }

  private async _addDateField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="DateTime" DisplayName="${fieldName}" Name="${fieldName}" Required="${required ? 'TRUE' : 'FALSE'}" Format="DateOnly" />`;
    await this._addField(listTitle, xml);
  }

  private async _addChoiceField(
    listTitle: string,
    fieldName: string,
    choices: string[],
    required: boolean,
    defaultValue?: string
  ): Promise<void> {
    const choiceXml = choices.map(c => `<CHOICE>${c}</CHOICE>`).join('');
    const defaultXml = defaultValue ? `<Default>${defaultValue}</Default>` : '';
    const xml = `<Field Type="Choice" DisplayName="${fieldName}" Name="${fieldName}" Required="${required ? 'TRUE' : 'FALSE'}" Format="Dropdown">${defaultXml}<CHOICES>${choiceXml}</CHOICES></Field>`;
    await this._addField(listTitle, xml);
  }

  private async _addPersonField(listTitle: string, fieldName: string): Promise<void> {
    const xml = `<Field Type="User" DisplayName="${fieldName}" Name="${fieldName}" Required="FALSE" UserSelectionMode="PeopleOnly" UserSelectionScope="0" />`;
    await this._addField(listTitle, xml);
  }

  private async _addLookupField(
    listTitle: string,
    fieldName: string,
    lookupListId: string,
    showField: string
  ): Promise<void> {
    const webId = await this._getWebId();
    const xml = `<Field Type="Lookup" DisplayName="${fieldName}" Name="${fieldName}" Required="FALSE" List="{${lookupListId}}" WebId="${webId}" ShowField="${showField}" />`;
    await this._addField(listTitle, xml);
  }

  private async _addCalculatedField(
    listTitle: string,
    fieldName: string,
    formula: string,
    resultType: string
  ): Promise<void> {
    const xml = `<Field Type="Calculated" DisplayName="${fieldName}" Name="${fieldName}" ResultType="${resultType}" ReadOnly="TRUE"><Formula>${formula}</Formula></Field>`;
    await this._addField(listTitle, xml);
  }

  private async _addBooleanField(listTitle: string, fieldName: string, defaultValue: boolean): Promise<void> {
    const xml = `<Field Type="Boolean" DisplayName="${fieldName}" Name="${fieldName}"><Default>${defaultValue ? '1' : '0'}</Default></Field>`;
    await this._addField(listTitle, xml);
  }

  private async _getWebId(): Promise<string> {
    const url = `${this._siteUrl}/_api/web?$select=Id`;
    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    return data.Id;
  }
}
