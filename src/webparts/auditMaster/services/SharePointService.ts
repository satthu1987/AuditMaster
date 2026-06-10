import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  LIST_NAMES,
  IAuditMasterItem,
  ICategoryItem,
  IISOClauseItem,
  ISegmentServiceItem,
  IUserRoleItem
} from '../models';

/**
 * Central service for all SharePoint CRUD operations against the five
 * Audit Master lists. Field names match the exact SharePoint list structure.
 */
export class SharePointService {
  private _context: WebPartContext;
  private _siteUrl: string;

  constructor(context: WebPartContext, siteUrl: string) {
    this._context = context;
    this._siteUrl = siteUrl;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Generic helpers
  // ═══════════════════════════════════════════════════════════════════════════

  private _headers(method: 'GET' | 'POST' | 'MERGE' | 'DELETE'): ISPHttpClientOptions {
    const h: any = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': ''
    };
    if (method === 'MERGE' || method === 'DELETE') {
      h['IF-MATCH'] = '*';
      h['X-HTTP-Method'] = method;
    }
    return { headers: h } as ISPHttpClientOptions;
  }

  private _listUrl(listTitle: string): string {
    return `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;
  }

  private async _get<T>(url: string): Promise<T> {
    const response: SPHttpClientResponse = await this._context.spHttpClient.get(
      url, SPHttpClient.configurations.v1, this._headers('GET')
    );
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`GET ${url} failed: ${response.status} – ${text}`);
    }
    return response.json() as Promise<T>;
  }

  private async _post<T>(url: string, body: any): Promise<T> {
    const options: ISPHttpClientOptions = {
      ...this._headers('POST'),
      body: JSON.stringify(body)
    };
    const response: SPHttpClientResponse = await this._context.spHttpClient.post(
      url, SPHttpClient.configurations.v1, options
    );
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`POST ${url} failed: ${response.status} – ${text}`);
    }
    return response.json() as Promise<T>;
  }

  private async _merge(url: string, body: any): Promise<void> {
    const options: ISPHttpClientOptions = {
      ...this._headers('MERGE'),
      body: JSON.stringify(body)
    };
    const response: SPHttpClientResponse = await this._context.spHttpClient.post(
      url, SPHttpClient.configurations.v1, options
    );
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`MERGE ${url} failed: ${response.status} – ${text}`);
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Current User
  // ═══════════════════════════════════════════════════════════════════════════

  public async getCurrentUserId(): Promise<number> {
    const data = await this._get<any>(`${this._siteUrl}/_api/web/currentuser?$select=Id`);
    return data.Id;
  }

  public async getCurrentUser(): Promise<{ Id: number; Title: string; Email: string; LoginName: string }> {
    const data = await this._get<any>(
      `${this._siteUrl}/_api/web/currentuser?$select=Id,Title,Email,LoginName`
    );
    return { Id: data.Id, Title: data.Title, Email: data.Email, LoginName: data.LoginName };
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Category List
  // ═══════════════════════════════════════════════════════════════════════════

  public async getCategories(): Promise<ICategoryItem[]> {
    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.CATEGORY)}/items?$select=Id,Category,Notes,CategoryGrouping1&$orderby=Category asc&$top=500`
    );
    return data.value;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  ISO Clause List
  // ═══════════════════════════════════════════════════════════════════════════

  public async getISOClauses(): Promise<IISOClauseItem[]> {
    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.ISO_CLAUSE)}/items?$select=Id,ISOClause,ISOlevel2,ISOlevel1,Article,Order0&$orderby=Order0 asc&$top=500`
    );
    return data.value;
  }

  /**
   * Fetch a single ISOClause list item by its Id, including all the fields
   * needed for cascading auto-population (ISOClause, ISOlevel2 → ISO Chapter,
   * ISOlevel1 → ISO Level, Article).
   */
  public async getISOClauseById(itemId: number): Promise<IISOClauseItem> {
    const data = await this._get<IISOClauseItem>(
      `${this._listUrl(LIST_NAMES.ISO_CLAUSE)}/items(${itemId})?$select=Id,ISOClause,ISOlevel2,ISOlevel1,Article,Order0`
    );
    return data;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Segment_Service List
  // ═══════════════════════════════════════════════════════════════════════════

  public async getSegmentServices(): Promise<ISegmentServiceItem[]> {
    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.SEGMENT_SERVICE)}/items?$select=Id,Title,Service,Service_Short,Division,Order0,DetailCategory&$orderby=Order0 asc&$top=500`
    );
    return data.value;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  UserRole List
  // ═══════════════════════════════════════════════════════════════════════════

  public async getUserRoles(): Promise<IUserRoleItem[]> {
    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.USER_ROLE)}/items?$select=Id,Title,AccountId,Role&$top=500`
    );
    return data.value;
  }

  public async getUserRoleForUser(userId: number): Promise<IUserRoleItem | null> {
    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.USER_ROLE)}/items?$select=Id,Title,AccountId,Role&$filter=AccountId eq ${userId}&$top=1`
    );
    return data.value.length > 0 ? data.value[0] : null;
  }

  /**
   * Fetch users from the UserRole list filtered by a specific role.
   * Expands the Account (Person) field to get Title and EMail for display.
   */
  public async getUsersByRole(role: string): Promise<IUserRoleItem[]> {
    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.USER_ROLE)}/items` +
      `?$select=Id,Title,AccountId,Role,Account/Id,Account/Title,Account/EMail` +
      `&$expand=Account` +
      `&$filter=Role eq '${role}'` +
      `&$top=500`
    );
    return data.value;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Audit Master List – READ
  // ═══════════════════════════════════════════════════════════════════════════

  private _serviceFieldInternalName: string = 'Service';
  private _serviceFieldEntityName: string = 'Service';
  private _serviceFieldType: string = 'Text';
  private _serviceFieldResolved: boolean = false;

  private _isServiceLookup(): boolean {
    return this._serviceFieldType === 'Lookup' || this._serviceFieldType === 'LookupMulti';
  }

  private async _ensureAuditServiceFieldNames(): Promise<void> {
    if (this._serviceFieldResolved) {
      return;
    }

    try {
      const data = await this._get<{ value: Array<{ Title?: string; InternalName?: string; EntityPropertyName?: string; TypeAsString?: string }> }>(
        `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/fields?$select=Title,InternalName,EntityPropertyName,TypeAsString`
      );
      const serviceField = data.value.find(f => (f.Title || '').toLowerCase() === 'service');
      if (serviceField) {
        this._serviceFieldInternalName = serviceField.InternalName || 'Service';
        this._serviceFieldEntityName = (serviceField.EntityPropertyName || serviceField.InternalName || 'Service').replace(/Id$/, '');
        this._serviceFieldType = serviceField.TypeAsString || 'Text';
      }
    } catch (e) {
      console.warn('[SharePointService] Could not resolve Service field metadata, using defaults.', e);
    }

    this._serviceFieldResolved = true;
  }

  private _auditSelectFields(): string {
    const serviceSelectName = this._isServiceLookup()
      ? `${this._serviceFieldEntityName}Id`
      : this._serviceFieldEntityName;

    return [
      'Id', 'Title',
      'PICId', 'QualityManagerId', 'AuditorId', 'VerifierId',
      'Status', 'FindingType', 'Region',
      'VerificationResult', 'InternalExternal',
      'QLVerification',
      'FindingDescription', 'QuickFix', 'ActionTaken',
      'RCDescription', 'PIONumber',
      'DueDate', 'AuditDate', 'VerificationDate', 'ClosedDate',
      'EvidenceLink',
      'ISOClauseId', serviceSelectName,
      'CategoryId',
      'Article',
      'Segment',
      'Year', 'FindingNumber', 'VerificationDateCalculated',
      'Created', 'Modified'
    ].join(',');
  }

  private _auditExpandFields(): string {
    const fields: string[] = [
      'PIC', 'QualityManager', 'Auditor', 'Verifier',
      'ISOClause',
      'FindingISOChapter', 'Category',
      'ISOClauseforArticle_lookup',
      'Service_Lookup'
    ];
    if (this._isServiceLookup()) {
      fields.splice(6, 0, this._serviceFieldEntityName);
    }
    return fields.join(',');
  }

  private _auditExpandSelect(): string {
    const fields: string[] = [
      'PIC/Id', 'PIC/Title', 'PIC/EMail',
      'QualityManager/Id', 'QualityManager/Title', 'QualityManager/EMail',
      'Auditor/Id', 'Auditor/Title', 'Auditor/EMail',
      'Verifier/Id', 'Verifier/Title', 'Verifier/EMail',
      'ISOClause/Id', 'ISOClause/ISOClause',
      'ISOClause/ISOlevel1', 'ISOClause/ISOlevel2',
      'ISOClause/Article',
      'FindingISOChapter/Id', 'FindingISOChapter/ISOClause',
      'Category/Id', 'Category/Category',
      'ISOClauseforArticle_lookup/Id', 'ISOClauseforArticle_lookup/ISOClause',
      'Service_Lookup/Id', 'Service_Lookup/Title'
    ];
    if (this._isServiceLookup()) {
      fields.splice(16, 0, `${this._serviceFieldEntityName}/Id`, `${this._serviceFieldEntityName}/Service`);
    }
    return fields.join(',');
  }

  // ───────────────────────────────────────────────────────────────────────────
  //  Lookup-expansion helpers
  //
  //  SharePoint lookup columns are NOT returned as objects by default. To get
  //  the related list's display columns (e.g. ISO clause text, Service name,
  //  Category) the REST query MUST:
  //    1. $expand the lookup field by its internal name, and
  //    2. $select the specific sub-columns (Lookup/DisplayColumn).
  //  If even ONE expanded field name is wrong, SharePoint rejects the whole
  //  request with HTTP 400 and every field comes back blank. To stay resilient
  //  we attempt the fully-expanded query first and, on failure, fall back to a
  //  base query (no $expand) so the form still loads with the raw FK ids while
  //  logging a clear diagnostic about which expand likely failed.
  // ───────────────────────────────────────────────────────────────────────────

  /** Build the query-string for the fully-expanded audit item read. */
  private _auditExpandedQuery(): string {
    return (
      `?$select=${this._auditSelectFields()},${this._auditExpandSelect()}` +
      `&$expand=${this._auditExpandFields()}`
    );
  }

  /** Build the query-string for the base (no-lookup-expansion) fallback read. */
  private _auditBaseQuery(): string {
    return `?$select=${this._auditSelectFields()}`;
  }

  /**
   * Log the lookup data that came back so the actual SharePoint field/property
   * names and values are visible in the browser console. This makes it obvious
   * whether a lookup (e.g. ISO clause) was expanded correctly or is missing.
   */
  private _debugLogAuditLookups(context: string, items: IAuditMasterItem[]): void {
    try {
      console.log(`[SharePointService] ${context}: received ${items.length} item(s).`);
      if (items.length === 0) { return; }
      const sample = items[0];
      console.log(`[SharePointService] ${context}: sample item property names →`, Object.keys(sample));
      console.log(`[SharePointService] ${context}: sample lookup values →`, {
        ISOClauseId: sample.ISOClauseId,
        ISOClause: sample.ISOClause,
        ServiceId: sample.ServiceId,
        Service: sample.Service,
        Segment: sample.Segment,
        CategoryId: sample.CategoryId,
        Category: sample.Category,
        Article: sample.Article,
        Division: sample.Division
      });
    } catch (e) {
      // Never let logging break a data fetch.
      console.warn(`[SharePointService] ${context}: debug logging failed`, e);
    }
  }

  /**
   * Fetch a list of audit items, trying the fully-expanded query first and
   * gracefully falling back to a base (no-expand) query if the expand fails.
   */
  private async _fetchAuditItems(
    context: string,
    filter: string
  ): Promise<IAuditMasterItem[]> {
    await this._ensureAuditServiceFieldNames();
    const base = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items`;
    const tail = `${filter}&$orderby=Modified desc&$top=500`;
    const expandedUrl = `${base}${this._auditExpandedQuery()}${tail}`;

    try {
      console.log(`[SharePointService] ${context}: GET (expanded) ${expandedUrl}`);
      const data = await this._get<{ value: IAuditMasterItem[] }>(expandedUrl);
      this._debugLogAuditLookups(context, data.value);
      return data.value;
    } catch (err) {
      console.warn(
        `[SharePointService] ${context}: expanded query failed – falling back to ` +
        `a non-expanded read. This usually means a lookup field internal name in ` +
        `_auditExpandFields/_auditExpandSelect does not match the SharePoint list. ` +
        `Lookup display values may be blank until the names are corrected.`,
        err
      );
      const fallbackUrl = `${base}${this._auditBaseQuery()}${tail}`;
      console.log(`[SharePointService] ${context}: GET (fallback) ${fallbackUrl}`);
      const data = await this._get<{ value: IAuditMasterItem[] }>(fallbackUrl);
      this._debugLogAuditLookups(`${context} (fallback)`, data.value);
      return data.value;
    }
  }

  /**
   * Get all Audit Master items (Admin view).
   */
  public async getAllAuditItems(): Promise<IAuditMasterItem[]> {
    return this._fetchAuditItems('getAllAuditItems', '');
  }

  /**
   * Get Audit items assigned to a specific PIC.
   */
  public async getAuditItemsByPIC(userId: number): Promise<IAuditMasterItem[]> {
    return this._fetchAuditItems('getAuditItemsByPIC', `&$filter=PICId eq ${userId}`);
  }

  /**
   * Get Audit items pending verification filtered by Service lookup IDs
   * (Verifier view).
   */
  public async getAuditItemsForVerifier(
    segmentServiceIds: number[]
  ): Promise<IAuditMasterItem[]> {
    await this._ensureAuditServiceFieldNames();
    let filter = '';
    if (segmentServiceIds.length > 0) {
      if (this._isServiceLookup()) {
        const conditions = segmentServiceIds.map(id => `${this._serviceFieldEntityName}Id eq ${id}`);
        filter = `&$filter=(${conditions.join(' or ')}) and (Status eq 'Pending Verification')`;
      } else {
        const segmentServices = await this.getSegmentServices();
        const serviceNames = segmentServices
          .filter(s => segmentServiceIds.indexOf(s.Id) !== -1)
          .map(s => (s.Service || s.Title || '').trim())
          .filter(s => s.length > 0);
        if (serviceNames.length > 0) {
          const conditions = serviceNames.map(name => `${this._serviceFieldEntityName} eq '${name.replace(/'/g, "''")}'`);
          filter = `&$filter=(${conditions.join(' or ')}) and (Status eq 'Pending Verification')`;
        } else {
          filter = `&$filter=Status eq 'Pending Verification'`;
        }
      }
    } else {
      filter = `&$filter=Status eq 'Pending Verification'`;
    }
    return this._fetchAuditItems('getAuditItemsForVerifier', filter);
  }

  /**
   * Get a single Audit Master item by ID, with all lookup fields expanded.
   * Falls back to a non-expanded read if the expand fails so the form can still
   * open. The returned object is logged so the actual lookup values/property
   * names are visible for debugging "blank lookup" issues.
   */
  public async getAuditItemById(itemId: number): Promise<IAuditMasterItem> {
    await this._ensureAuditServiceFieldNames();
    const base = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items(${itemId})`;
    const expandedUrl = `${base}${this._auditExpandedQuery()}`;

    try {
      console.log(`[SharePointService] getAuditItemById(${itemId}): GET (expanded) ${expandedUrl}`);
      const item = await this._get<IAuditMasterItem>(expandedUrl);
      this._debugLogAuditLookups(`getAuditItemById(${itemId})`, [item]);
      return item;
    } catch (err) {
      console.warn(
        `[SharePointService] getAuditItemById(${itemId}): expanded query failed – ` +
        `falling back to a non-expanded read. Verify the lookup field internal ` +
        `names in _auditExpandFields/_auditExpandSelect against the SharePoint list.`,
        err
      );
      const fallbackUrl = `${base}${this._auditBaseQuery()}`;
      console.log(`[SharePointService] getAuditItemById(${itemId}): GET (fallback) ${fallbackUrl}`);
      const item = await this._get<IAuditMasterItem>(fallbackUrl);
      this._debugLogAuditLookups(`getAuditItemById(${itemId}) (fallback)`, [item]);
      return item;
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Audit Master List – CREATE / UPDATE
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Create a new Audit Master item. Returns the created item.
   */
  public async createAuditItem(item: Partial<IAuditMasterItem>): Promise<IAuditMasterItem> {
    await this._ensureAuditServiceFieldNames();
    const payload = this._buildAuditPayload(item);
    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items`;
    return this._post<IAuditMasterItem>(url, payload);
  }

  /**
   * Update an existing Audit Master item.
   */
  public async updateAuditItem(itemId: number, item: Partial<IAuditMasterItem>): Promise<void> {
    await this._ensureAuditServiceFieldNames();
    const payload = this._buildAuditPayload(item);
    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items(${itemId})`;
    await this._merge(url, payload);
  }

  /**
   * Build a clean payload stripping read-only, calculated, and expand fields.
   * Uses the exact internal field names that SharePoint REST API expects.
   */
  private _buildAuditPayload(item: Partial<IAuditMasterItem>): any {
    const payload: any = {};

    // Single line of text
    if (item.Title !== undefined) payload.Title = item.Title;
    if (item.PIONumber !== undefined) payload.PIONumber = item.PIONumber;

    // Choice fields
    if (item.Status !== undefined) payload.Status = item.Status;
    if (item.FindingType !== undefined) payload.FindingType = item.FindingType;
    if (item.VerificationResult !== undefined) payload.VerificationResult = item.VerificationResult;
    if (item.QLVerification !== undefined) payload.Q_x0026_Lverification = item.QLVerification;

    // Multiple lines of text
    if (item.FindingDescription !== undefined) payload.FindingDescription = item.FindingDescription;
    if (item.QuickFix !== undefined) payload.Quickfix = item.QuickFix;
    if (item.ActionTaken !== undefined) payload.ActionTaken = item.ActionTaken;
    if (item.RCDescription !== undefined) payload.RCdescription = item.RCDescription;

    // Date fields
    if (item.DueDate !== undefined) payload.DueDate = item.DueDate;
    if (item.AuditDate !== undefined) payload.AuditDate = item.AuditDate;
    if (item.VerificationDate !== undefined) payload.VerificationDate = item.VerificationDate;

    // Hyperlink fields – stored as { Url, Description }
    if (item.EvidenceLink !== undefined) {
      payload.EvidenceLink = item.EvidenceLink;
    }

    // Person field IDs
    if (item.PICId !== undefined) payload.PICId = item.PICId;
    if (item.AuditorId !== undefined) payload.AuditorId = item.AuditorId;

    // Lookup IDs (primary lookups only – dependent lookups are read-only)
    if (item.ISOClauseId !== undefined) payload.ISOclauseId = item.ISOClauseId;
    if (item.CategoryId !== undefined) payload.CategoryId = item.CategoryId;

    // Service supports both Text and Lookup list designs
    if (item.Service !== undefined && !this._isServiceLookup()) {
      payload[this._serviceFieldInternalName] = item.Service;
    }
    if (item.ServiceId !== undefined && this._isServiceLookup()) {
      payload[`${this._serviceFieldEntityName}Id`] = item.ServiceId;
    }

    // NOTE: Dependent lookups (#8 Segment, #31 Article, #33 Division) are read-only.
    // NOTE: Calculated fields (#25 Year, #26 FindingNumber, #34 VerificationDateCalculated) are read-only.

    return payload;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  People Picker helper
  // ═══════════════════════════════════════════════════════════════════════════

  public async ensureUser(loginName: string): Promise<number> {
    const url = `${this._siteUrl}/_api/web/ensureuser`;
    const body = { logonName: loginName };
    const options: ISPHttpClientOptions = {
      ...this._headers('POST'),
      body: JSON.stringify(body)
    };
    const response: SPHttpClientResponse = await this._context.spHttpClient.post(
      url, SPHttpClient.configurations.v1, options
    );
    if (!response.ok) {
      throw new Error(`ensureUser failed for ${loginName}`);
    }
    const data = await response.json();
    return data.Id;
  }

  public async searchPeople(query: string): Promise<Array<{ Id: number; Title: string; Email: string }>> {
    const url = `${this._siteUrl}/_api/web/siteusers?$select=Id,Title,Email&$filter=substringof('${encodeURIComponent(query)}',Title)&$top=10`;
    const data = await this._get<{ value: any[] }>(url);
    return data.value;
  }
}
