import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  LIST_NAMES,
  IAuditMasterItem,
  ICategoryItem,
  IISOClauseItem,
  ISegmentServiceItem,
  IUserRoleItem,
  AuditType
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
  private _segmentFieldInternalName: string = 'Segment';
  private _segmentFieldType: string = 'Text';
  private _divisionFieldInternalName: string = 'Division';
  private _divisionFieldType: string = 'Text';
  private _isoClauseFieldEntityName: string = 'ISOClause';
  private _isoChapterFieldEntityName: string = 'ISOChapter';
  private _isoChapterFieldInternalName: string = 'ISOChapter';
  private _isoChapterFieldType: string = 'Lookup';
  private _qlVerificationFieldEntityName: string = 'QLVerification';
  private _findingIdFieldEntityName: string = 'FindingId';
  private _auditTypeFieldEntityName: string = 'AuditType';
  private _verifierFieldEntityName: string = 'Verifier';
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

      const segmentField = data.value.find(f => (f.Title || '').toLowerCase() === 'segment');
      if (segmentField) {
        this._segmentFieldInternalName = segmentField.InternalName || 'Segment';
        this._segmentFieldType = segmentField.TypeAsString || 'Text';
      }

      const divisionField = data.value.find(f => (f.Title || '').toLowerCase() === 'division');
      if (divisionField) {
        this._divisionFieldInternalName = divisionField.InternalName || 'Division';
        this._divisionFieldType = divisionField.TypeAsString || 'Text';
      }

      const normalize = (value?: string): string => (value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      const isWritableField = (field: { TypeAsString?: string }): boolean => {
        const type = normalize(field.TypeAsString);
        return type !== 'calculated' && type !== 'computed' && type !== 'counter';
      };

      const qlNames = new Set(['qlverification', 'qx0026lverification']);
      const isoNames = new Set(['isoclause']);
      const isoChapterNames = new Set(['isochapter', 'findingisochapter', 'findingx0020isox0020chapter']);
      const findingIdNames = new Set(['findingid', 'findingx0020id']);
      const findingLegacyNames = new Set(['findingnumber']);
      const auditTypeNames = new Set(['audittype', 'internalexternal', 'internalx002fexternal']);
      const verifierNames = new Set(['verifier', 'verifiedby', 'verifiedx0020by']);

      const qlVerificationField = data.value.find(f => {
        const title = normalize(f.Title);
        const internal = normalize(f.InternalName);
        const entity = normalize(f.EntityPropertyName);
        return qlNames.has(title) || qlNames.has(internal) || qlNames.has(entity);
      });
      if (qlVerificationField) {
        this._qlVerificationFieldEntityName = qlVerificationField.EntityPropertyName || qlVerificationField.InternalName || 'QLVerification';
      }

      const isoClauseField = data.value.find(f => {
        const title = normalize(f.Title);
        const internal = normalize(f.InternalName);
        const entity = normalize(f.EntityPropertyName);
        return isoNames.has(title) || isoNames.has(internal) || isoNames.has(entity);
      });
      if (isoClauseField) {
        this._isoClauseFieldEntityName = (isoClauseField.EntityPropertyName || isoClauseField.InternalName || 'ISOClause').replace(/Id$/, '');
      }

      let isoChapterField = data.value.find(f => {
        if (!isWritableField(f)) {
          return false;
        }
        const title = normalize(f.Title);
        const internal = normalize(f.InternalName);
        const entity = normalize(f.EntityPropertyName);
        return isoChapterNames.has(title) || isoChapterNames.has(internal) || isoChapterNames.has(entity);
      });
      if (!isoChapterField) {
        isoChapterField = data.value.find(f => {
          if (!isWritableField(f)) {
            return false;
          }
          const type = normalize(f.TypeAsString);
          const title = normalize(f.Title);
          const internal = normalize(f.InternalName);
          const entity = normalize(f.EntityPropertyName);
          const isLookup = type === 'lookup' || type === 'lookupmulti';
          const hasIso = title.indexOf('iso') !== -1 || internal.indexOf('iso') !== -1 || entity.indexOf('iso') !== -1;
          const hasChapter = title.indexOf('chapter') !== -1 || internal.indexOf('chapter') !== -1 || entity.indexOf('chapter') !== -1;
          return isLookup && hasIso && hasChapter;
        });
      }
      if (isoChapterField) {
        this._isoChapterFieldEntityName = (isoChapterField.EntityPropertyName || isoChapterField.InternalName || 'ISOChapter').replace(/Id$/, '');
        this._isoChapterFieldInternalName = isoChapterField.InternalName || 'ISOChapter';
        this._isoChapterFieldType = isoChapterField.TypeAsString || 'Lookup';
      }

      const findingIdField = data.value.find(f => {
        if (!isWritableField(f)) {
          return false;
        }
        const title = normalize(f.Title);
        const internal = normalize(f.InternalName);
        const entity = normalize(f.EntityPropertyName);
        return findingIdNames.has(title) || findingIdNames.has(internal) || findingIdNames.has(entity);
      });
      if (findingIdField) {
        this._findingIdFieldEntityName = findingIdField.InternalName || findingIdField.EntityPropertyName || 'FindingId';
      } else {
        const legacyCalculatedField = data.value.find(f => {
          const title = normalize(f.Title);
          const internal = normalize(f.InternalName);
          const entity = normalize(f.EntityPropertyName);
          return findingLegacyNames.has(title) || findingLegacyNames.has(internal) || findingLegacyNames.has(entity);
        });
        if (legacyCalculatedField) {
          console.warn('[SharePointService] FindingNumber is calculated and not writable. Add/confirm a writable Finding ID text column.');
        }
      }

      const auditTypeField = data.value.find(f => {
        if (!isWritableField(f)) {
          return false;
        }
        const title = normalize(f.Title);
        const internal = normalize(f.InternalName);
        const entity = normalize(f.EntityPropertyName);
        return auditTypeNames.has(title) || auditTypeNames.has(internal) || auditTypeNames.has(entity);
      });
      if (auditTypeField) {
        this._auditTypeFieldEntityName = auditTypeField.EntityPropertyName || auditTypeField.InternalName || 'AuditType';
      }

      const verifierField = data.value.find(f => {
        if (!isWritableField(f)) {
          return false;
        }
        const title = normalize(f.Title);
        const internal = normalize(f.InternalName);
        const entity = normalize(f.EntityPropertyName);
        return verifierNames.has(title) || verifierNames.has(internal) || verifierNames.has(entity);
      });
      if (verifierField) {
        this._verifierFieldEntityName = (verifierField.EntityPropertyName || verifierField.InternalName || 'Verifier').replace(/Id$/, '');
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
      this._auditTypeFieldEntityName,
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
      'Year', this._findingIdFieldEntityName, 'VerificationDateCalculated',
      'Created', 'Modified'
    ].join(',');
  }

  private _auditExpandFields(): string {
    const fields: string[] = [
      'PIC', 'QualityManager', 'Auditor', 'Verifier',
      'ISOClause', 'Category'
    ];

    if (this._isServiceLookup()) {
      fields.push(this._serviceFieldEntityName);
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
      'ISOClause/ISOlevel1', 'ISOClause/ISOlevel2', 'ISOClause/Article',
      'Category/Id', 'Category/Category'
    ];

    if (this._isServiceLookup()) {
      fields.push(`${this._serviceFieldEntityName}/Id`, `${this._serviceFieldEntityName}/Service`);
    }

    return fields.join(',');
  }

  private _auditTypePrefix(auditType?: string): string {
    if (!auditType) {
      return 'CUA';
    }

    if (auditType === AuditType.Internal) {
      return 'INA';
    }

    if (auditType === AuditType.External) {
      return 'EXA';
    }

    return 'CUA';
  }

  private _pad3(value: number): string {
    if (value < 10) {
      return `00${value}`;
    }
    if (value < 100) {
      return `0${value}`;
    }
    return `${value}`;
  }

  public async generateFindingId(auditType: AuditType, auditDate?: string): Promise<string> {
    await this._ensureAuditServiceFieldNames();

    const now = auditDate ? new Date(auditDate) : new Date();
    const year = now.getFullYear();
    const yy = `${year}`.slice(-2);
    const prefix = this._auditTypePrefix(auditType);

    const auditTypeField = this._auditTypeFieldEntityName;
    const findingIdField = this._findingIdFieldEntityName;

    const select = `Id,AuditDate,Created,${auditTypeField},${findingIdField}`;
    const filter = `${auditTypeField} eq '${String(auditType).replace(/'/g, "''")}'`;
    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items?$select=${select}&$filter=${filter}&$top=5000`;
    const data = await this._get<{ value: any[] }>(url);

    let maxSeq = 0;

    for (const item of data.value) {
      const dt = item.AuditDate || item.Created;
      if (!dt) {
        continue;
      }
      const d = new Date(dt);
      if (d.getFullYear() !== year) {
        continue;
      }

      const findingId = typeof item[findingIdField] === 'string' ? item[findingIdField] : '';
      const expectedPrefix = `${prefix}-${yy}-`;
      if (findingId.indexOf(expectedPrefix) !== 0) {
        continue;
      }
      const seqPart = findingId.slice(expectedPrefix.length);
      if (!/^\d{3}$/.test(seqPart)) {
        continue;
      }
      const seq = parseInt(seqPart, 10);
      if (seq > maxSeq) {
        maxSeq = seq;
      }
    }

    const nextSeq = maxSeq + 1;
    return `${prefix}-${yy}-${this._pad3(nextSeq)}`;
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
    console.log('[SharePointService] createAuditItem: payload →', payload);

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
    if (item.FindingId !== undefined) payload[this._findingIdFieldEntityName] = item.FindingId;

    // Choice fields
    if (item.Status !== undefined) payload.Status = item.Status;
    if (item.FindingType !== undefined) payload.FindingType = item.FindingType;
    if (item.AuditType !== undefined) payload[this._auditTypeFieldEntityName] = item.AuditType;
    if (item.VerificationResult !== undefined) payload.VerificationResult = item.VerificationResult;
    if (item.QLVerification !== undefined) payload[this._qlVerificationFieldEntityName] = item.QLVerification;

    // Multiple lines of text
    if (item.FindingDescription !== undefined) payload.FindingDescription = item.FindingDescription;
    if (item.QuickFix !== undefined) payload.QuickFix = item.QuickFix;
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
    if (item.VerifierId !== undefined) payload[`${this._verifierFieldEntityName}Id`] = item.VerifierId;

    // Lookup IDs (primary lookups only – dependent lookups are read-only)
    if (item.ISOClauseId !== undefined) payload[`${this._isoClauseFieldEntityName}Id`] = item.ISOClauseId;
    if (item.ISOChapterId !== undefined) {
      if (this._isoChapterFieldType === 'Lookup' || this._isoChapterFieldType === 'LookupMulti') {
        payload[`${this._isoChapterFieldEntityName}Id`] = item.ISOChapterId;
      } else {
        payload[this._isoChapterFieldInternalName] = item.ISOChapterId;
      }
    } else if (item.ISOClauseId !== undefined) {
      if (this._isoChapterFieldType === 'Lookup' || this._isoChapterFieldType === 'LookupMulti') {
        payload[`${this._isoChapterFieldEntityName}Id`] = item.ISOClauseId;
      } else {
        payload[this._isoChapterFieldInternalName] = item.ISOClauseId;
      }
    }
    if (item.CategoryId !== undefined) payload.CategoryId = item.CategoryId;

    // Service supports both Text and Lookup list designs
    if (item.Service !== undefined && !this._isServiceLookup()) {
      payload[this._serviceFieldInternalName] = item.Service;
    }
    if (item.ServiceId !== undefined && this._isServiceLookup()) {
      payload[`${this._serviceFieldEntityName}Id`] = item.ServiceId;
    }

    // Segment/Division can be either text or lookup depending on list design
    if (item.Segment !== undefined) {
      payload[this._segmentFieldInternalName] = item.Segment;
    }
    if (item.Division !== undefined) {
      payload[this._divisionFieldInternalName] = item.Division;
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
