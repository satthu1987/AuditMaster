import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  LIST_NAMES,
  DivisionChoices,
  DetailCategoryChoices,
  AuditStatus,
  FindingType,
  VerificationResult,
  UserRole,
  QLVerification
} from '../models';

/**
 * Service responsible for ensuring all required SharePoint lists exist.
 * Creates lists and their fields if they do not already exist on the site.
 *
 * Handles primary lookups AND dependent (secondary) lookups:
 *  - Service → Segment_Service.Service  (primary)
 *    ├── Segment  → Segment_Service.Title    (dependent)
 *    └── Division → Segment_Service.Division (dependent)
 *  - ISO clause → ISOClause.ISOClause (primary)
 *    └── Article  → ISOClause.Article  (dependent)
 *  - Category → AuditCategory.Category (primary)
 *  - Finding ISO Chapter → ISOClause.ISOClause (separate primary lookup)
 */
export class ListProvisioningService {
  private _context: WebPartContext;
  private _siteUrl: string;
  private _webId: string | null = null;

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

    // Cache web ID once for lookup fields
    this._webId = await this._getWebId();

    // Reference lists first (they are lookup targets)
    await this._provisionCategoryList();
    await this._provisionISOClauseList();
    await this._provisionSegmentServiceList();
    await this._provisionUserRoleList();

    // Main list last (depends on reference lists for lookups)
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

    // Title is built-in (used as Segment name)
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
  // Fields match the exact structure from the SharePoint list screenshot:
  //
  //  #  Field Name                    Type
  //  ── ──────────────────────────── ──────────────────────────────────
  //  1  Title                         Single line of text (built-in)
  //  2  PIC                           Person or Group
  //  3  Status                        Choice
  //  4  FindingType                   Choice
  //  5  Region                        Choice
  //  6  ISO clause                    Lookup → ISOClause.ISOClause
  //  7  Service                       Lookup → Segment_Service.Service
  //  8  Segment                       Dependent lookup (from Service → Title)
  //  9  Quality Manager               Person or Group
  // 10  Auditor                       Person or Group
  // 11  Finding Description           Multiple lines of text
  // 12  Required RCA                  Yes/No
  // 13  Quick fix                     Multiple lines of text
  // 14  Action Taken                  Multiple lines of text
  // 15  DueDate                       Date and Time
  // 16  Verification Result           Choice
  // 17  Internal/External             Choice
  // 18  Finding ISO Chapter           Lookup → ISOClause.ISOClause
  // 19  PIONumber                     Single line of text
  // 20  RC description                Multiple lines of text
  // 21  Audit Date                    Date and Time
  // 22  Verified by                   Person or Group
  // 23  VerificationDate              Date and Time
  // 24  Evidence (link)               Hyperlink or Picture
  // 25  Year                          Calculated
  // 26  FindingNumber                 Calculated
  // 27  Closed Date                   Date and Time
  // 28  Confluence Page               Hyperlink or Picture
  // 29  Category                      Lookup → AuditCategory.Category
  // 30  Article                       Dependent lookup (from ISO clause → Article)
  // 31  Division                      Dependent lookup (from Service → Division)
  // 32  VerificationDateCalculated    Calculated
  // 33  Action Status                 Choice
  // 34  Q&L verification              Choice

  private async _provisionAuditMasterList(): Promise<void> {
    const listTitle = LIST_NAMES.AUDIT_MASTER;
    if (await this._listExists(listTitle)) {
      console.log(`[ListProvisioning] List "${listTitle}" already exists.`);
      return;
    }

    await this._createList(listTitle, 'Main audit findings list', 100);

    // ── Person or Group fields ──────────────────────────────────────────────
    // 2. PIC
    await this._addPersonField(listTitle, 'PIC');
    // 9. Quality Manager
    //await this._addPersonField(listTitle, 'Quality Manager');
    // 10. Auditor
    await this._addPersonField(listTitle, 'Auditor');
    // 22. Verified by
    await this._addPersonField(listTitle, 'Verified by');

    // ── Choice fields ───────────────────────────────────────────────────────
    // 3. Status
    await this._addChoiceField(
      listTitle, 'Status',
      Object.values(AuditStatus),
      true, AuditStatus.Open
    );
    // 4. FindingType
    await this._addChoiceField(
      listTitle, 'FindingType',
      Object.values(FindingType),
      false
    );
    // 5. Region
   /*  await this._addChoiceField(
      listTitle, 'Region',
      RegionChoices,
      false
    ); */
    // 16. Verification Result
    await this._addChoiceField(
      listTitle, 'Verification Result',
      Object.values(VerificationResult),
      false, VerificationResult.No
    );
    // 17. Internal/External
    /* await this._addChoiceField(
      listTitle, 'Internal/External',
      Object.values(InternalExternal),
      false
    ); */
    // 34. Q&L verification
    await this._addChoiceField(
      listTitle,
      'Q&L verification',
      Object.values(QLVerification),
      false,
      QLVerification.No
    );

    // ── Multiple lines of text fields ───────────────────────────────────────
    // 11. Finding Description
    await this._addNoteField(listTitle, 'Finding Description', false);
    // 13. Quick fix
    await this._addNoteField(listTitle, 'Quick fix', false);
    // 14. Action Taken
    await this._addNoteField(listTitle, 'Action Taken', false);
    // 20. RC description
    await this._addNoteField(listTitle, 'RC description', false);

    // ── Yes/No field ────────────────────────────────────────────────────────
    // 12. Required RCA
   // await this._addBooleanField(listTitle, 'Required RCA', false);

    // ── Date and Time fields ────────────────────────────────────────────────
    // 15. DueDate
    await this._addDateField(listTitle, 'DueDate', false);
    // 21. Audit Date
    await this._addDateField(listTitle, 'Audit Date', false);
    // 23. VerificationDate
    await this._addDateField(listTitle, 'VerificationDate', false);
    // 27. Closed Date
    //await this._addDateField(listTitle, 'Closed Date', false);

    // ── Single line of text fields ──────────────────────────────────────────
    // 19. PIONumber
    await this._addTextField(listTitle, 'PIONumber', false);

    // ── Hyperlink or Picture fields ─────────────────────────────────────────
    // 24. Evidence (link)
    await this._addUrlField(listTitle, 'Evidence (link)');
    // 28. Confluence Page
    await this._addUrlField(listTitle, 'Confluence Page');

    // ── Primary Lookup fields ───────────────────────────────────────────────

    // 6. ISO clause → ISOClause list, showing ISOClause field
    const isoClauseListId = await this._getListId(LIST_NAMES.ISO_CLAUSE);
    let isoClauseLookupFieldId: string | null = null;
    if (isoClauseListId) {
      isoClauseLookupFieldId = await this._addLookupField(
        listTitle, 'ISO clause', isoClauseListId, 'ISOClause'
      );
    }

    // 7. Service → Segment_Service list, showing Service field
    const segmentServiceListId = await this._getListId(LIST_NAMES.SEGMENT_SERVICE);
    let serviceLookupFieldId: string | null = null;
    if (segmentServiceListId) {
      serviceLookupFieldId = await this._addLookupField(
        listTitle, 'Service', segmentServiceListId, 'Service'
      );
    }

    // 18. Finding ISO Chapter → ISOClause list, showing ISOClause field (separate lookup)
    if (isoClauseListId) {
      await this._addLookupField(
        listTitle, 'Finding ISO Chapter', isoClauseListId, 'ISOClause'
      );
    }

    // 29. Category → AuditCategory list, showing Category field
    const categoryListId = await this._getListId(LIST_NAMES.CATEGORY);
    if (categoryListId) {
      await this._addLookupField(
        listTitle, 'Category', categoryListId, 'Category'
      );
    }

    // ── Dependent (secondary) Lookup fields ─────────────────────────────────
    // These reference the same list as the primary lookup but show a different
    // column. They use the primary lookup's field ID as their FieldRef.

    // 8. Segment – dependent lookup from Service → Segment_Service.Title
    if (segmentServiceListId && serviceLookupFieldId) {
      await this._addDependentLookupField(
        listTitle, 'Segment', segmentServiceListId, 'Title', serviceLookupFieldId
      );
    }

    // 30. Article – dependent lookup from ISO clause → ISOClause.Article
    //     (labelled "ISOClause for Article_lookup" in SP)
    if (isoClauseListId && isoClauseLookupFieldId) {
      await this._addDependentLookupField(
        listTitle, 'Article', isoClauseListId, 'Article', isoClauseLookupFieldId
      );
    }

    // 31. Division – dependent lookup from Service → Segment_Service.Division
    //     (labelled "Service_Lookup: Division look up" in SP)
    if (segmentServiceListId && serviceLookupFieldId) {
      await this._addDependentLookupField(
        listTitle, 'Division', segmentServiceListId, 'Division', serviceLookupFieldId
      );
    }

    // ── Calculated fields ───────────────────────────────────────────────────
    // 25. Year – extracts year from Audit Date
    await this._addCalculatedField(
      listTitle,
      'Year',
      '=TEXT([Audit Date],"YYYY")',
      'Text',
      ['Audit Date']
    );

    // 26. FindingNumber – auto-generated finding reference number
    //     Format: combines Year + ID for a unique finding number
    await this._addCalculatedField(
      listTitle,
      'FindingNumber',
      '=TEXT([Audit Date],"YYYY")&"-"&TEXT([ID],"0000")',
      'Text',
      ['Audit Date', 'ID']
    );

    // 32. VerificationDateCalculated – shows VerificationDate or fallback
    await this._addCalculatedField(
      listTitle,
      'VerificationDateCalculated',
      '=IF(ISBLANK([VerificationDate]),"",[VerificationDate])',
      'DateTime',
      ['VerificationDate']
    );

    console.log(`[ListProvisioning] List "${listTitle}" created with all 34 fields.`);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Helper Methods
  // ═══════════════════════════════════════════════════════════════════════════

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

  /**
   * Generic field creation via CAML SchemaXml.
   */
  private async _addField(listTitle: string, fieldXml: string): Promise<string | null> {
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
      return null;
    }

    // Return the created field's Id for dependent lookup chaining
    try {
      const data = await response.json();
      return data.d ? data.d.Id : (data.Id || null);
    } catch {
      return null;
    }
  }

  // ── Typed field helpers ───────────────────────────────────────────────────

  private async _addTextField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="Text" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="${required ? 'TRUE' : 'FALSE'}" MaxLength="255" />`;
    await this._addField(listTitle, xml);
  }

  private async _addNoteField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="Note" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="${required ? 'TRUE' : 'FALSE'}" NumLines="6" RichText="TRUE" RichTextMode="FullHtml" />`;
    await this._addField(listTitle, xml);
  }

  private async _addNumberField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="Number" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="${required ? 'TRUE' : 'FALSE'}" Decimals="0" Min="0" />`;
    await this._addField(listTitle, xml);
  }

  private async _addDateField(listTitle: string, fieldName: string, required: boolean): Promise<void> {
    const xml = `<Field Type="DateTime" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="${required ? 'TRUE' : 'FALSE'}" Format="DateOnly" />`;
    await this._addField(listTitle, xml);
  }

  private async _addChoiceField(
    listTitle: string,
    fieldName: string,
    choices: string[],
    required: boolean,
    defaultValue?: string
  ): Promise<void> {
    const choiceXml = choices.map(c => `<CHOICE>${this._escapeXml(c)}</CHOICE>`).join('');
    const defaultXml = defaultValue ? `<Default>${this._escapeXml(defaultValue)}</Default>` : '';
    const xml = `<Field Type="Choice" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="${required ? 'TRUE' : 'FALSE'}" Format="Dropdown">${defaultXml}<CHOICES>${choiceXml}</CHOICES></Field>`;
    await this._addField(listTitle, xml);
  }

  private async _addPersonField(listTitle: string, fieldName: string): Promise<void> {
    const xml = `<Field Type="User" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="FALSE" UserSelectionMode="PeopleOnly" UserSelectionScope="0" />`;
    await this._addField(listTitle, xml);
  }

  private async _addBooleanField(listTitle: string, fieldName: string, defaultValue: boolean): Promise<void> {
    const xml = `<Field Type="Boolean" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}"><Default>${defaultValue ? '1' : '0'}</Default></Field>`;
    await this._addField(listTitle, xml);
  }

  private async _addUrlField(listTitle: string, fieldName: string): Promise<void> {
    const xml = `<Field Type="URL" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="FALSE" Format="Hyperlink" />`;
    await this._addField(listTitle, xml);
  }

  /**
   * Create a primary lookup field. Returns the created field's ID (GUID)
   * so dependent lookups can reference it.
   */
  private async _addLookupField(
    listTitle: string,
    fieldName: string,
    lookupListId: string,
    showField: string
  ): Promise<string | null> {
    const webId = this._webId || await this._getWebId();
    const xml = `<Field Type="Lookup" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="FALSE" List="{${lookupListId}}" WebId="${webId}" ShowField="${showField}" />`;
    return this._addField(listTitle, xml);
  }

  /**
   * Create a dependent (secondary) lookup field.
   * A dependent lookup shows a different column from the same lookup list
   * as a primary lookup. It references the primary lookup field via FieldRef.
   *
   * Uses the REST API endpoint for adding dependent lookup fields.
   */
  private async _addDependentLookupField(
    listTitle: string,
    fieldName: string,
    lookupListId: string,
    showField: string,
    primaryLookupFieldId: string
  ): Promise<void> {
    const webId = this._webId || await this._getWebId();

    // Use the AddDependentLookupField REST endpoint
    const listApiUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;
    const url = `${listApiUrl}/fields/adddependentlookupfield(displayname='${encodeURIComponent(fieldName)}',primarylookupfieldid='${primaryLookupFieldId}',showfield='${showField}')`;

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': ''
      }
    };

    const response = await this._context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!response.ok) {
      // Fallback: try creating as a regular lookup with FieldRef attribute
      console.warn(`[ListProvisioning] Dependent lookup via REST failed for "${fieldName}", trying CAML fallback...`);
      const xml = `<Field Type="Lookup" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" Required="FALSE" List="{${lookupListId}}" WebId="${webId}" ShowField="${showField}" FieldRef="${primaryLookupFieldId}" ReadOnly="TRUE" UnlimitedLengthInDocumentLibrary="FALSE" />`;
      await this._addField(listTitle, xml);
    } else {
      console.log(`[ListProvisioning] Dependent lookup "${fieldName}" created successfully.`);
    }
  }

  /**
   * Create a calculated field with field references for the formula.
   */
  private async _addCalculatedField(
    listTitle: string,
    fieldName: string,
    formula: string,
    resultType: string,
    fieldRefs: string[]
  ): Promise<void> {
    const fieldRefsXml = fieldRefs
      .map(ref => `<FieldRef Name="${this._toInternalName(ref)}" />`)
      .join('');
    const xml = `<Field Type="Calculated" DisplayName="${this._escapeXml(fieldName)}" Name="${this._toInternalName(fieldName)}" ResultType="${resultType}" ReadOnly="TRUE"><Formula>${this._escapeXml(formula)}</Formula><FieldRefs>${fieldRefsXml}</FieldRefs></Field>`;
    await this._addField(listTitle, xml);
  }

  // ── Utility helpers ───────────────────────────────────────────────────────

  private async _getWebId(): Promise<string> {
    const url = `${this._siteUrl}/_api/web?$select=Id`;
    const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const data = await response.json();
    return data.Id;
  }

  /**
   * Convert a display name to a safe internal field name.
   * SharePoint internal names cannot contain spaces or most special characters.
   */
  private _toInternalName(displayName: string): string {
    return displayName
      .replace(/[^a-zA-Z0-9_]/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '');
  }

  /**
   * Escape special XML characters in attribute values and text content.
   */
  private _escapeXml(str: string): string {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
