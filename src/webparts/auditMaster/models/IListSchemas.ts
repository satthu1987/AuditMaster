/**
 * SharePoint List Schemas for Audit Master
 * Defines the structure of all five SharePoint lists used by the system.
 */

// ─── Category List ───────────────────────────────────────────────────────────
export interface ICategoryItem {
    Id: number;
    Category: string;
    Notes: string;
    CategoryGrouping1: string;
  }
  
  // ─── ISO Clause List ─────────────────────────────────────────────────────────
  export interface IISOClauseItem {
    Id: number;
    ISOClause: string;
    ISOlevel2: string;
    ISOlevel1: string;
    Article: string;
    Order: number;
  }
  
  // ─── Segment_Service List ────────────────────────────────────────────────────
  export interface ISegmentServiceItem {
    Id: number;
    Title: string;
    Service: string;
    Service_Short: string;
    Division: string;       // Choice field
    Order: number;
    DetailCategory: string; // Choice field
  }
  
  // ─── UserRole List ───────────────────────────────────────────────────────────
  export interface IUserRoleItem {
    Id: number;
    Title: string;
    AccountId: number;
    AccountStringId?: string;
    Account?: {
      Id: number;
      Title: string;
      EMail: string;
      LoginName?: string;
    };
    Role: UserRole;
  }
  
  export enum UserRole {
    Admin = 'Admin',
    PIC = 'PIC',
    Verifier = 'Verifier'
  }
  
  // ─── Audit Master List ───────────────────────────────────────────────────────
  export interface IAuditMasterItem {
    Id: number;
  
    // Basic text fields
    Title: string;
    AuditId: string;
    Description: string;
    AuditYear: string;
    Finding: string;
    RootCause: string;
    CorrectiveAction: string;
    PreventiveAction: string;
    Evidence: string;
    Remarks: string;
  
    // Choice fields
    Status: AuditStatus;
    Priority: AuditPriority;
    AuditType: AuditType;
    Source: AuditSource;
    VerificationResult: VerificationResult;
  
    // Lookup fields (Category)
    CategoryId: number;
    Category?: ICategoryItem;
  
    // Lookup fields (ISO Clause)
    ISOClauseId: number;
    ISOClause?: IISOClauseItem;
  
    // Lookup fields (Segment/Service)
    SegmentServiceId: number;
    SegmentService?: ISegmentServiceItem;
  
    // Person or Group fields
    PICId: number;
    PIC?: { Id: number; Title: string; EMail: string };
    VerifierId: number;
    Verifier?: { Id: number; Title: string; EMail: string };
    AuditorId: number;
    Auditor?: { Id: number; Title: string; EMail: string };
    AuditeeId: number;
    Auditee?: { Id: number; Title: string; EMail: string };
    CreatedById: number;
    CreatedByUser?: { Id: number; Title: string; EMail: string };
  
    // Date fields
    AuditDate: string;       // ISO date string
    DueDate: string;
    CompletionDate: string;
    VerificationDate: string;
    TargetCloseDate: string;
    Created: string;
    Modified: string;
  
    // Number fields
    ExtensionCount: number;
    DaysOpen: number;         // Calculated
  
    // Boolean fields
    IsOverdue: boolean;       // Calculated
    RequiresFollowUp: boolean;
  
    // Multi-line / Rich text
    ActionPlan: string;
    FollowUpNotes: string;
  
    // Additional metadata
    Department: string;
    Location: string;
    RiskRating: RiskRating;
  
    // Attachments flag
    AttachmentFiles?: any[];
  }
  
  // ─── Choice Enums ────────────────────────────────────────────────────────────
  export enum AuditStatus {
    Open = 'Open',
    InProgress = 'In Progress',
    PendingVerification = 'Pending Verification',
    Closed = 'Closed',
    Overdue = 'Overdue',
    Cancelled = 'Cancelled'
  }
  
  export enum AuditPriority {
    Critical = 'Critical',
    High = 'High',
    Medium = 'Medium',
    Low = 'Low'
  }
  
  export enum AuditType {
    Internal = 'Internal',
    External = 'External',
    Surveillance = 'Surveillance',
    Certification = 'Certification',
    Recertification = 'Recertification'
  }
  
  export enum AuditSource {
    ManagementReview = 'Management Review',
    InternalAudit = 'Internal Audit',
    ExternalAudit = 'External Audit',
    CustomerComplaint = 'Customer Complaint',
    RegulatoryInspection = 'Regulatory Inspection',
    SelfAssessment = 'Self-Assessment'
  }
  
  export enum VerificationResult {
    NotVerified = 'Not Verified',
    Effective = 'Effective',
    NotEffective = 'Not Effective',
    PartiallyEffective = 'Partially Effective'
  }
  
  export enum RiskRating {
    VeryHigh = 'Very High',
    High = 'High',
    Medium = 'Medium',
    Low = 'Low',
    VeryLow = 'Very Low'
  }
  
  // ─── Division choices for Segment_Service ────────────────────────────────────
  export const DivisionChoices: string[] = [
    'Operations',
    'Finance',
    'Human Resources',
    'Information Technology',
    'Quality Assurance',
    'Engineering',
    'Sales & Marketing',
    'Legal & Compliance',
    'Supply Chain',
    'Research & Development'
  ];
  
  // ─── Detail Category choices for Segment_Service ─────────────────────────────
  export const DetailCategoryChoices: string[] = [
    'Core Service',
    'Support Service',
    'Administrative',
    'Technical',
    'Consulting',
    'Outsourced'
  ];
  
  // ─── List Name Constants ─────────────────────────────────────────────────────
  export const LIST_NAMES = {
    CATEGORY: 'AuditCategory',
    ISO_CLAUSE: 'ISOClause',
    SEGMENT_SERVICE: 'Segment_Service',
    USER_ROLE: 'UserRole',
    AUDIT_MASTER: 'AuditMaster'
  } as const;
  