export interface DocxportInitOptions {
  assetBaseUrl?: string | URL;
  diagnosticTracing?: boolean;
  environment?: string;
}

export type FieldMode = "none" | "evaluate" | "cache";
export type TrackedChangeMode = "accept" | "reject" | "inline" | "split";
export type HeaderFooterSelection = "none" | "first" | "last";
export type ExportPreset = "rich" | "plain";
export type MathOutputFormat = "none" | "mathml" | "latex" | "unicodemath" | "text";
export type MathDelimiterStyle = "dollar" | "backslash" | "auto";

export type ExportPhase = "opening" | "preparing" | "converting" | "finalizing" | "completed";

export interface ExportProgress {
  phase: ExportPhase;
  completedUnits: number;
  totalUnits: number;
  percentage: number | null;
}

export interface ExportProgressOptions {
  onProgress?: (progress: ExportProgress) => void;
}

export interface FieldOptions {
  mode?: FieldMode;
  variables?: Record<string, string | null>;
}

export interface HtmlOptions {
  mathOutputFormat?: MathOutputFormat;
  emitImages?: boolean;
  emitParagraphMetadata?: boolean;
  emitStyleFont?: boolean;
  emitRunColor?: boolean;
  emitRunBackground?: boolean;
  emitTableBorders?: boolean;
  emitDocumentColors?: boolean;
  emitParagraphAlignment?: boolean;
  preserveListSymbols?: boolean;
  richTables?: boolean;
  emitSectionHeadersFooters?: boolean;
  emitUnreferencedBookmarks?: boolean;
  emitPageNumbers?: boolean;
  emitFieldInstructions?: boolean;
  usePlainComments?: boolean;
  emitCustomProperties?: boolean;
  emitTimeline?: boolean;
  stylesheetHref?: string;
  embedDefaultStylesheet?: boolean;
  rootCssClass?: string;
  trackedChangeMode?: TrackedChangeMode;
  headerSelection?: HeaderFooterSelection;
  footerSelection?: HeaderFooterSelection;
}

export interface MarkdownOptions {
  mathOutputFormat?: MathOutputFormat;
  emitMathDelimiters?: boolean;
  mathDelimiterStyle?: MathDelimiterStyle;
  emitImages?: boolean;
  emitStyleFont?: boolean;
  emitRunColor?: boolean;
  emitRunBackground?: boolean;
  emitTableBorders?: boolean;
  emitDocumentColors?: boolean;
  emitParagraphAlignment?: boolean;
  emitRichLayoutHtml?: boolean;
  preserveListSymbols?: boolean;
  richTables?: boolean;
  usePlainCodeBlocks?: boolean;
  useMarkdownInlineStyles?: boolean;
  emitSectionHeadersFooters?: boolean;
  emitUnreferencedBookmarks?: boolean;
  emitPageNumbers?: boolean;
  emitFieldInstructions?: boolean;
  usePlainComments?: boolean;
  emitCustomProperties?: boolean;
  emitTimeline?: boolean;
  trackedChangeMode?: TrackedChangeMode;
}

export interface TextOptions {
  mathOutputFormat?: MathOutputFormat;
  trackedChangeMode?: "accept" | "reject";
  imagePlaceholder?: string;
  emitDocumentProperties?: boolean;
  emitCustomProperties?: boolean;
}

export type ExportRequest = (
  | { format: "html"; preset?: ExportPreset; fields?: FieldOptions; html?: HtmlOptions }
  | { format: "markdown"; preset?: ExportPreset; fields?: FieldOptions; markdown?: MarkdownOptions }
  | { format: "text"; fields?: FieldOptions; text?: TextOptions }
) & ExportProgressOptions;

export interface PdfExportRequest extends ExportProgressOptions {
  fields?: FieldOptions;
  /** Family name previously supplied to registerPdfFont; defaults to the first registered family. */
  pdfFontFamily?: string;
}

export interface ResolveRequest { fields?: FieldOptions; }
export interface CoreProperties {
  title: string | null;
  subject: string | null;
  creator: string | null;
  lastModifiedBy: string | null;
  revision: string | null;
  created: string | null;
  modified: string | null;
  description: string | null;
  category: string | null;
  keywords: string | null;
}
export interface ExtendedProperties {
  application: string | null;
  applicationVersion: string | null;
  template: string | null;
  pages: string | null;
  words: string | null;
  characters: string | null;
  lines: string | null;
  paragraphs: string | null;
  totalTime: string | null;
}
export interface LanguageRatio { code: string; ratio: number; }
export interface DocumentInfo {
  coreProperties: CoreProperties;
  extendedProperties: ExtendedProperties | null;
  language: LanguageRatio[] | null;
  hasTrackedChanges: boolean;
  hasComments: boolean;
}

export interface DocEmailAddress { address: string; displayName?: string; }
export interface DocEmailAttachment { fileName: string; content: Uint8Array | ArrayBuffer; }
export interface DocEmailEnvelope {
  subject?: string;
  introduction?: string;
  to?: DocEmailAddress[];
  cc?: DocEmailAddress[];
  bcc?: DocEmailAddress[];
  replyTo?: DocEmailAddress[];
  attachments?: DocEmailAttachment[];
  importance?: "low" | "normal" | "high";
  sensitivity?: "normal" | "personal" | "private" | "confidential";
  requestDeliveryReceipt?: boolean;
  requestReadReceipt?: boolean;
  visible?: boolean;
  categories?: string;
  deliverAfter?: string;
  expiresAt?: string;
}
export type DocEnvelopeEdit =
  | { operation: "set"; envelope: DocEmailEnvelope }
  | { operation: "visibility"; visible: boolean }
  | { operation: "remove" };

export interface Docxport {
  convertOmml(omml: string, format?: "mathml" | "html" | "latex" | "unicodemath" | "text"): Promise<string>;
  /** Inspect metadata, revisions, and comments in DOCX or binary DOC bytes. */
  inspect(input: Uint8Array | ArrayBuffer): Promise<DocumentInfo>;
  /** Export DOCX or binary DOC bytes. Binary DOC is projected to basic DOCX first. */
  export(input: Uint8Array | ArrayBuffer, request: ExportRequest): Promise<string>;
  /** Register OpenType font bytes before exporting PDF. */
  registerPdfFont(family: string, fontBytes: Uint8Array | ArrayBuffer): void;
  /** Register a specific bold/italic face for a family. */
  registerPdfFontFace(family: string, bold: boolean, italic: boolean, fontBytes: Uint8Array | ArrayBuffer): void;
  /** List font families referenced by DOCX runs, styles, and theme, including binary DOC projections. */
  inspectPdfFontFamilies(input: Uint8Array | ArrayBuffer): Promise<string[]>;
  /** Render a DOC or DOCX into PDF bytes using registered fonts. */
  exportPdf(input: Uint8Array | ArrayBuffer, request?: PdfExportRequest): Promise<Uint8Array>;
  /** Directly project binary DOC bytes into a basic DOCX, without a DOCX visitor pass. */
  projectDocx(input: Uint8Array | ArrayBuffer): Promise<Uint8Array>;
  /** Walk DOCX or binary DOC input and write a plain text binary DOC. */
  exportDoc(input: Uint8Array | ArrayBuffer, request?: ResolveRequest): Promise<Uint8Array>;
  /** Queue envelope edits and return a newly saved binary DOC. */
  editDocEnvelope(input: Uint8Array | ArrayBuffer, edits: DocEnvelopeEdit[]): Promise<Uint8Array>;
  /** Return resolved DOCX bytes; binary DOC input is projected first. */
  resolveDocx(input: Uint8Array | ArrayBuffer, request?: ResolveRequest): Promise<Uint8Array>;
}

export function createDocxport(options?: DocxportInitOptions): Promise<Docxport>;
