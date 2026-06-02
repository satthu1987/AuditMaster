/**
 * SharePoint List Schemas for Audit Master
 * Defines the structure of all five SharePoint lists used by the system.
 * Matches the EXACT 36-field structure from the SharePoint Audit Master list screenshot.
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
  Title: string;          // Segment name – dependent lookup target "Segment"
  Service: string;        // Primary lookup target for "Service" in Audit Master
  Service_Short: string;
  Division: string;       // Choice field – dependent lookup target for "Division"
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

// ─── Audit Master List (36 fields) ──────────────────────────────────────────
// Matches the EXACT field structure from the SharePoint list screenshot.
//
//  #  Field Name                         Type
//  ── ─────────────────────────────────  ────────────────────────────────────
//  1  Title                              Single line of text (built-in)
//  2  PIC                                Person or Group
//  3  Status                             Choice
//  4  FindingType                        Choice
//  5  Region                             Choice
//  6  ISO clause                         Lookup → ISOClause.ISOClause
//  7  Service                            Lookup → Segment_Service.Service
//  8  Segment                            Lookup (dependent from Service → Title)
//  9  Quality Manager                    Person or Group
// 10  Auditor                            Person or Group
// 11  Finding Description                Multiple lines of text
// 12  Required RCA                       Yes/No
// 13  Quick fix                          Multiple lines of text
// 14  Action Taken                       Multiple lines of text
// 15  DueDate                            Date and Time
// 16  Verification Result                Choice
// 17  Internal/External                  Choice
// 18  Finding ISO Chapter                Lookup → ISOClause.ISOClause
// 19  PIONumber                          Single line of text
// 20  RC description                     Multiple lines of text
// 21  Audit Date                         Date and Time
// 22  Verified by                        Person or Group
// 23  VerificationDate                   Date and Time
// 24  Evidence (link)                    Hyperlink or Picture
// 25  Year                               Calculated
// 26  FindingNumber                      Calculated
// 27  Closed Date                        Date and Time
// 28  Confluence Page                    Hyperlink or Picture
// 29  Category                           Lookup → AuditCategory.Category
// 30  ISOClause for Article_lookup       Lookup (primary, → ISOClause list)
// 31  Article                            Lookup (dependent from #30 → Article)
// 32  Service_Lookup                     Lookup (primary, → Segment_Service list)
// 33  Service_Lookup: Division look up   Lookup (dependent from #32 → Division)
// 34  VerificationDateCalculated         Calculated
// 35  Action Status                      Choice
// 36  Q&L verification                   Choice

export interface IAuditMasterItem {
  Id: number;

  // 1. Title – Single line of text
  Title: string;

  // 2. PIC – Person or Group
  PICId: number;
  PIC?: { Id: number; Title: string; EMail: string };

  // 3. Status – Choice
  Status: AuditStatus;

  // 4. FindingType – Choice
  FindingType: FindingType;

  // 5. Region – Choice
  Region: string;

  // 6. ISO clause – Lookup (primary, to ISOClause list → ISOClause field)
  ISOClauseId: number;
  ISOClause?: { Id: number; ISOClause: string };

  // 7. Service – Lookup (primary, to Segment_Service list → Service field)
  ServiceId: number;
  Service?: { Id: number; Service: string };

  // 8. Segment – Dependent lookup from Service → Segment_Service.Title
  Segment?: string;

  // 9. Quality Manager – Person or Group
  QualityManagerId: number;
  QualityManager?: { Id: number; Title: string; EMail: string };

  // 10. Auditor – Person or Group
  AuditorId: number;
  Auditor?: { Id: number; Title: string; EMail: string };

  // 11. Finding Description – Multiple lines of text
  FindingDescription: string;

  // 12. Required RCA – Yes/No
  RequiredRCA: boolean;

  // 13. Quick fix – Multiple lines of text
  QuickFix: string;

  // 14. Action Taken – Multiple lines of text
  ActionTaken: string;

  // 15. DueDate – Date and Time
  DueDate: string;

  // 16. Verification Result – Choice
  VerificationResult: VerificationResult;

  // 17. Internal/External – Choice
  InternalExternal: InternalExternal;

  // 18. Finding ISO Chapter – Lookup (to ISOClause list → ISOClause field)
  FindingISOChapterId: number;
  FindingISOChapter?: { Id: number; ISOClause: string };

  // 19. PIONumber – Single line of text
  PIONumber: string;

  // 20. RC description – Multiple lines of text
  RCDescription: string;

  // 21. Audit Date – Date and Time
  AuditDate: string;

  // 22. Verified by – Person or Group
  VerifiedById: number;
  VerifiedBy?: { Id: number; Title: string; EMail: string };

  // 23. VerificationDate – Date and Time
  VerificationDate: string;

  // 24. Evidence (link) – Hyperlink or Picture
  EvidenceLink: IHyperlinkField;

  // 25. Year – Calculated (based on Audit Date)
  Year: string;

  // 26. FindingNumber – Calculated
  FindingNumber: string;

  // 27. Closed Date – Date and Time
  ClosedDate: string;

  // 28. Confluence Page – Hyperlink or Picture
  ConfluencePage: IHyperlinkField;

  // 29. Category – Lookup (to AuditCategory list → Category field)
  CategoryId: number;
  Category?: { Id: number; Category: string };

  // 31. Article – Dependent lookup from #30 (ISOClause for Article_lookup → Article)
  Article?: string;

  // 33. Service_Lookup: Division look up – Dependent lookup from #32 (Service_Lookup → Division)
  Service_Division?: string;

  // 34. VerificationDateCalculated – Calculated
  VerificationDateCalculated: string;

  // 35. Action Status – Choice
  ActionStatus: ActionStatus;

  // 36. Q&L verification – Choice
  QLVerification: QLVerification;

  // SharePoint system fields
  Created: string;
  Modified: string;
}

// ─── Hyperlink or Picture field shape ────────────────────────────────────────
export interface IHyperlinkField {
  Url: string;
  Description: string;
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

export enum FindingType {
  NC = 'NC',
  OBS = 'OBS',
  OFI = 'OFI'
}

export enum VerificationResult {
  Yes = 'Yes',
  No = 'No'
}

export enum InternalExternal {
  Internal = 'Internal',
  External = 'External'
}

export enum ActionStatus {
  Onprogress = 'On progress',
  Completed = 'Completed',
  ActionNotFilled = 'Action not filled',
  NoAction = 'No action',
  NeedVerification = 'Need verification',
  Closed = 'Closed',
  ProvideEvidence = 'Provide evidence'

}

export enum QLVerification {
  Yes = 'Yes',
  No = 'No'
}

// ─── Region choices ──────────────────────────────────────────────────────────
export const RegionChoices: string[] = [
  'Asia',
  'North America',
  'Europe',
  'Global'
];

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
