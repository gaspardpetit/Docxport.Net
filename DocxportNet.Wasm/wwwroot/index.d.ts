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

export interface ResolveRequest { fields?: FieldOptions; }
export interface DocumentInfo { hasTrackedChanges: boolean; }

export interface Docxport {
  convertOmml(omml: string, format?: "mathml" | "html" | "latex" | "unicodemath" | "text"): Promise<string>;
  inspect(input: Uint8Array | ArrayBuffer): Promise<DocumentInfo>;
  export(input: Uint8Array | ArrayBuffer, request: ExportRequest): Promise<string>;
  resolveDocx(input: Uint8Array | ArrayBuffer, request?: ResolveRequest): Promise<Uint8Array>;
}

export function createDocxport(options?: DocxportInitOptions): Promise<Docxport>;
