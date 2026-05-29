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
 * Audit Master lists.
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
    const data = await this._get<{ value: any[] }>(
      `${this._listUrl(LIST_NAMES.SEGMENT_SERVICE)}/items?$select=Id,Title,Service,Service_Short,Division,Order,DetailCategory&$orderby=Order asc&$top=500`
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

  // ═══════════════════════════════════════════════════════════════════════════
  //  Audit Master List – READ
  // ═══════════════════════════════════════════════════════════════════════════

  private readonly _auditSelectFields: string = [
    'Id', 'Title', 'AuditId', 'Description', 'AuditYear', 'Finding', 'RootCause',
    'CorrectiveAction', 'PreventiveAction', 'Evidence', 'Remarks', 'ActionPlan',
    'FollowUpNotes', 'Department', 'Location',
    'Status', 'Priority', 'AuditType', 'Source', 'VerificationResult', 'RiskRating',
    'CategoryId', 'ISOClauseId', 'SegmentServiceId',
    'PICId', 'VerifierId', 'AuditorId', 'AuditeeId', 'CreatedByUserId',
    'AuditDate', 'DueDate', 'CompletionDate', 'VerificationDate', 'TargetCloseDate',
    'ExtensionCount', 'DaysOpen', 'IsOverdue', 'RequiresFollowUp',
    'Created', 'Modified'
  ].join(',');

  private readonly _auditExpandFields: string = [
    'PIC', 'Verifier', 'Auditor', 'Auditee', 'CreatedByUser'
  ].join(',');

  /**
   * Get all Audit Master items (Admin view).
   */
  public async getAllAuditItems(): Promise<IAuditMasterItem[]> {
    const url =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items` +
      `?$select=${this._auditSelectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,Auditee/Id,Auditee/Title,Auditee/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}` +
      `&$orderby=Modified desc&$top=500`;

    const data = await this._get<{ value: IAuditMasterItem[] }>(url);
    return data.value;
  }

  /**
   * Get Audit items assigned to a specific PIC.
   */
  public async getAuditItemsByPIC(userId: number): Promise<IAuditMasterItem[]> {
    const url =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items` +
      `?$select=${this._auditSelectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,Auditee/Id,Auditee/Title,Auditee/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}` +
      `&$filter=PICId eq ${userId}` +
      `&$orderby=Modified desc&$top=500`;

    const data = await this._get<{ value: IAuditMasterItem[] }>(url);
    return data.value;
  }

  /**
   * Get Audit items pending verification filtered by Segment/Service IDs
   * (Verifier view – items that belong to their department/segment).
   */
  public async getAuditItemsForVerifier(
    segmentServiceIds: number[]
  ): Promise<IAuditMasterItem[]> {
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
      `?$select=${this._auditSelectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,Auditee/Id,Auditee/Title,Auditee/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}` +
      filter +
      `&$orderby=Modified desc&$top=500`;

    const data = await this._get<{ value: IAuditMasterItem[] }>(url);
    return data.value;
  }

  /**
   * Get a single Audit Master item by ID.
   */
  public async getAuditItemById(itemId: number): Promise<IAuditMasterItem> {
    const url =
      `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items(${itemId})` +
      `?$select=${this._auditSelectFields}` +
      `,PIC/Id,PIC/Title,PIC/EMail` +
      `,Verifier/Id,Verifier/Title,Verifier/EMail` +
      `,Auditor/Id,Auditor/Title,Auditor/EMail` +
      `,Auditee/Id,Auditee/Title,Auditee/EMail` +
      `,CreatedByUser/Id,CreatedByUser/Title,CreatedByUser/EMail` +
      `&$expand=${this._auditExpandFields}`;

    return this._get<IAuditMasterItem>(url);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  //  Audit Master List – CREATE / UPDATE
  // ═══════════════════════════════════════════════════════════════════════════

  /**
   * Create a new Audit Master item. Returns the created item.
   */
  public async createAuditItem(item: Partial<IAuditMasterItem>): Promise<IAuditMasterItem> {
    const payload = this._buildAuditPayload(item);
    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items`;
    return this._post<IAuditMasterItem>(url, payload);
  }

  /**
   * Update an existing Audit Master item.
   */
  public async updateAuditItem(itemId: number, item: Partial<IAuditMasterItem>): Promise<void> {
    const payload = this._buildAuditPayload(item);
    const url = `${this._listUrl(LIST_NAMES.AUDIT_MASTER)}/items(${itemId})`;
    await this._merge(url, payload);
  }

  /**
   * Build a clean payload from IAuditMasterItem partial, stripping
   * read-only and expand fields.
   */
  private _buildAuditPayload(item: Partial<IAuditMasterItem>): any {
    const payload: any = {};

    // Text fields
    const textFields = [
      'Title', 'AuditId', 'Description', 'AuditYear', 'Finding', 'RootCause',
      'CorrectiveAction', 'PreventiveAction', 'Evidence', 'Remarks',
      'ActionPlan', 'FollowUpNotes', 'Department', 'Location'
    ];
    for (const f of textFields) {
      if ((item as any)[f] !== undefined) payload[f] = (item as any)[f];
    }

    // Choice fields
    const choiceFields = [
      'Status', 'Priority', 'AuditType', 'Source', 'VerificationResult', 'RiskRating'
    ];
    for (const f of choiceFields) {
      if ((item as any)[f] !== undefined) payload[f] = (item as any)[f];
    }

    // Lookup ID fields
    const lookupFields = ['CategoryId', 'ISOClauseId', 'SegmentServiceId'];
    for (const f of lookupFields) {
      if ((item as any)[f] !== undefined) payload[f] = (item as any)[f];
    }

    // Person ID fields
    const personFields = ['PICId', 'VerifierId', 'AuditorId', 'AuditeeId', 'CreatedByUserId'];
    for (const f of personFields) {
      if ((item as any)[f] !== undefined) payload[f] = (item as any)[f];
    }

    // Date fields
    const dateFields = ['AuditDate', 'DueDate', 'CompletionDate', 'VerificationDate', 'TargetCloseDate'];
    for (const f of dateFields) {
      if ((item as any)[f] !== undefined) payload[f] = (item as any)[f];
    }

    // Number fields
    if (item.ExtensionCount !== undefined) payload.ExtensionCount = item.ExtensionCount;

    // Boolean fields
    if (item.RequiresFollowUp !== undefined) payload.RequiresFollowUp = item.RequiresFollowUp;

    return payload;
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
}
