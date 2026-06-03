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
import { AUDIT_FIELD_CONFIG } from '../components/config/AuditMasterSchema';
/**
 * Central service for all SharePoint CRUD operations against the five
 * Audit Master lists.
 */


interface IResolvedField {
  internalName: string;
  type: string;
}


export class SharePointService {
  private _auditFieldMap?: Record<string, IResolvedField>;
  private _context: WebPartContext;
  private _siteUrl: string;
  private _auditSelectFieldsCache?: string;
  private _segmentServicesCache?: ISegmentServiceItem[];
  private _auditMasterFieldsCache?: Array<{ Title: string; InternalName: string; TypeAsString: string }>;


  private _normalize(value: string): string {
    return String(value || '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');
  }

  constructor(context: WebPartContext, siteUrl: string) {
    this._context = context;
    this._siteUrl = siteUrl;
  }

  private async _getAuditFieldMap():
    Promise<Record<string, IResolvedField>> {

    if (this._auditFieldMap) {
      return this._auditFieldMap;
    }

    const fields =
      await this._getAuditMasterFields();

    const result:
      Record<string, IResolvedField> = {};

    for (const config of AUDIT_FIELD_CONFIG) {

      let match =
        fields.find(f =>
          config.types.includes(f.TypeAsString) &&
          config.aliases.some(alias =>
            this._normalize(f.Title) === alias ||
            this._normalize(f.InternalName) === alias
          )
        ) ||
        fields.find(f =>
          config.types.includes(f.TypeAsString) &&
          config.aliases.some(alias =>
            this._normalize(f.Title).includes(alias) ||
            this._normalize(f.InternalName).includes(alias)
          )
        );

      if (match) {
        result[config.key] = {
          internalName: match.InternalName,
          type: match.TypeAsString
        };
      }
    }

    this._auditFieldMap = result;

    console.log('Resolved fields', result);

    return result;
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

  /**
   * Generic GET returning JSON.
   */
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

  /**
   * Generic POST (create).
   */
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

  /**
   * Generic MERGE (update).
   */
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
      `${this._listUrl(LIST_NAMES.ISO_CLAUSE)}/items?$select=Id,ISOClause,ISOlevel2,ISOlevel1,Article,Order&$orderby=Order asc&$top=500`
    );
    return data.value;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Segment_Service List
  // ═══════════════════════════════════════════════════════════════════════════

  public async getSegmentServices(): Promise<ISegmentServiceItem[]> {
    if (this._segmentServicesCache) {
      return this._segmentServicesCache;
    }

    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.SEGMENT_SERVICE)}/items?$select=Id,Title,Service,Service_Short,Division,Order,DetailCategory&$orderby=Order asc&$top=500`
    );
    this._segmentServicesCache = data.value;
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

  // ═══════════════════════════════════════════════════════════════════════════
  //  Audit Master List – READ
  // ═══════════════════════════════════════════════════════════════════════════

  private readonly _auditSelectFieldsBase: string[] = [
    'Id', 'Title', 'AuditId',
    'Status', 'FindingType', 'Region', 'AuditType', 'VerificationResult', 'InternalExternal',
    'FindingDescription', 'RequiredRCA', 'QuickFix', 'ActionTaken', 'PIONumber', 'RCDescription',
    'CategoryId', 'ISOClauseId', 'Service', 'Segment', 'Article',
    'PICId', 'QualityManagerId', 'AuditorId', 'VerifierId', 'CreatedByUserId',
    'AuditDate', 'DueDate', 'ClosedDate', 'VerificationDate',
    'EvidenceLink', 'ConfluencePage',
    'Year', 'VerificationDateCalculated',
    'QLVerification',
    'Created', 'Modified'
  ];

  private readonly _auditExpandFields: string = [
    'PIC', 'Verifier', 'Auditor', 'CreatedByUser'
  ].join(',');

  private async _getAuditMasterFields(): Promise<Array<{ Title: string; InternalName: string; TypeAsString: string }>> {
    if (this._auditMasterFieldsCache) {
      return this._auditMasterFieldsCache;
    }

    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/fields?$select=Title,InternalName,TypeAsString&$top=500`;
    const data = await this._get<{ value: Array<{ Title: string; InternalName: string; TypeAsString: string }> }>(url);
    this._auditMasterFieldsCache = data.value;
    return data.value;
  }


  private async _getAuditSelectFields():
    Promise<string> {

    if (this._auditSelectFieldsCache) {
      return this._auditSelectFieldsCache;
    }

    const fieldMap =
      await this._getAuditFieldMap();

    const fields = [
      ...this._auditSelectFieldsBase
    ];

    for (const config of AUDIT_FIELD_CONFIG) {

      if (!config.includeInSelect) {
        continue;
      }

      const field =
        fieldMap[config.key];

      if (!field) {
        continue;
      }

      const selectField =
        ['Lookup', 'LookupMulti']
          .includes(field.type)
          ? `${field.internalName}Id`
          : field.internalName;

      if (!fields.includes(selectField)) {
        fields.push(selectField);
      }
    }

    this._auditSelectFieldsCache =
      fields.join(',');

    console.log(
      'Audit Select Fields',
      this._auditSelectFieldsCache
    );

    return this._auditSelectFieldsCache;
  }


  /**
   * Get all Audit Master items (Admin view).
   */
  public async getAllAuditItems(): Promise<IAuditMasterItem[]> {
    const selectFields = await this._getAuditSelectFields();
    const url =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items` +
      `?$select=${selectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}` +
      `&$orderby=Modified desc&$top=500`;

    const data = await this._get<{ value: any[] }>(url);
    return data.value;
  }

  /**
   * Get Audit items assigned to a specific PIC.
   */
  public async getAuditItemsByPIC(userId: number): Promise<IAuditMasterItem[]> {
    const selectFields = await this._getAuditSelectFields();
    const url =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items` +
      `?$select=${selectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}` +
      `&$filter=PICId eq ${userId}` +
      `&$orderby=Modified desc&$top=500`;

    const data = await this._get<{ value: any[] }>(url);
    return data.value;
  }

  /**
   * Get Audit items pending verification filtered by Segment/Service IDs
   * (Verifier view – items that belong to their department/segment).
   */
  public async getAuditItemsForVerifier(
    segmentServiceIds: number[]
  ): Promise<IAuditMasterItem[]> {
    const selectFields = await this._getAuditSelectFields();
    // Build filter for segment service ids
    let filter = '';
    if (segmentServiceIds.length > 0) {
      const conditions = segmentServiceIds.map(id => `SegmentServiceId eq ${id}`);
      filter = `&$filter=(${conditions.join(' or ')}) and (Status eq 'Pending Verification')`;
    } else {
      filter = `&$filter=Status eq 'Pending Verification'`;
    }

    const url =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items` +
      `?$select=${selectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,Auditee/Id,Auditee/Title,Auditee/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}` +
      filter +
      `&$orderby=Modified desc&$top=500`;

    const data = await this._get<{ value: any[] }>(url);
    return data.value;
  }

  /**
   * Get a single Audit Master item by ID.
   */
  public async getAuditItemById(itemId: number): Promise<IAuditMasterItem> {
    const selectFields = await this._getAuditSelectFields();
    const url =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items(${itemId})` +
      `?$select=${selectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,Auditee/Id,Auditee/Title,Auditee/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}`;

    const data = await this._get<any>(url);
    return data;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Audit Master List – CREATE / UPDATE
  // ═══════════════════════════════════════════════════════════════════════════

  private _toPositiveId(value: unknown): number | undefined {
    if (value === undefined || value === null) {
      return undefined;
    }
    if (typeof value === 'string' && value.trim() === '') {
      return undefined;
    }
    const parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed <= 0) {
      return undefined;
    }
    return Math.floor(parsed);
  }

  private async _toAuditWritePayload(item: Partial<IAuditMasterItem>): Promise<any> {
    const payload: any = { ...item };

    const userFieldTargets = AUDIT_FIELD_CONFIG
      .filter(config =>
        config.types.some(t =>
          ['User', 'UserMulti'].includes(t)
        )
      )
      .map(config => ({
        name: config.key,
        target: `${config.key}Id`
      }));

    userFieldTargets.forEach((field) => {
      const objectValue = (payload as any)[field.name];
      const explicitId = (payload as any)[field.target];
      const idCandidate = explicitId ?? objectValue?.Id;
      const idValue = this._toPositiveId(idCandidate);
      delete (payload as any)[field.name];
      if (idValue !== undefined) {
        (payload as any)[field.target] = idValue;
      } else {
        delete (payload as any)[field.target];
      }
    });
    return payload;
  }

  /**
   * Create a new Audit Master item. Returns the created item.
   */
  public async createAuditItem(item: Partial<IAuditMasterItem>): Promise<IAuditMasterItem> {
    const payload: any = await this._toAuditWritePayload(item);

    const currentUserId = await this.getCurrentUserId();
    payload.CreatedByUserId = currentUserId;
    delete payload.AuditId;
    console.log('Creating item with payload', payload);
    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items`;
    const created = await this._post<IAuditMasterItem>(url, payload);

    if (!created.Id) {
      throw new Error('Failed to create audit item. No SharePoint item ID returned.');
    }
    const generatedAuditId =
      `AUD-${new Date().getFullYear()}-${String(created.Id).padStart(4, '0')}`;
    const updateUrl =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items(${created.Id})`;

    await this._merge(updateUrl, {
      AuditId: generatedAuditId
    });
    return {
      ...created,
      AuditId: generatedAuditId
    };
  }

  /**
   * Update an existing Audit Master item.
   */
  public async updateAuditItem(itemId: number, item: Partial<IAuditMasterItem>): Promise<void> {
    const payload: any = await this._toAuditWritePayload(item);
    const currentUserId = await this.getCurrentUserId();
    payload.CreatedByUserId = currentUserId;
    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items(${itemId})`;
    await this._merge(url, payload);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  People Picker helper
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Resolve a user by email or login name and return their SP user ID.
   */
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

  /**
   * Search people by query string (for people picker autocomplete).
   */
  public async searchPeople(query: string): Promise<Array<{ Id: number; Title: string; Email: string }>> {
    const url = `${this._siteUrl}/_api/web/siteusers?$select=Id,Title,Email&$filter=substringof('${encodeURIComponent(query)}',Title)&$top=10`;
    const data = await this._get<{ value: any[] }>(url);
    return data.value;
  }

  public async getUsersByIds(userIds: number[]): Promise<Array<{ Id: number; Title: string; Email: string }>> {
    if (userIds.length === 0) {
      return [];
    }

    const uniqueIds = Array.from(new Set(userIds));
    const idFilter = uniqueIds.map(id => `Id eq ${id}`).join(' or ');
    const url = `${this._siteUrl}/_api/web/siteusers?$select=Id,Title,Email&$filter=${encodeURIComponent(idFilter)}&$top=500`;
    const data = await this._get<{ value: Array<{ Id: number; Title: string; Email: string }> }>(url);
    return data.value;
  }
}
