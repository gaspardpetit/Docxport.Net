export interface DocxportInitOptions {
  assetBaseUrl?: string | URL;
  diagnosticTracing?: boolean;
  environment?: string;
}

export type FieldMode = "none" | "evaluate" | "cache";
export type TrackedChangeMode = "accept" | "reject" | "inline" | "split";
export type HeaderFooterSelection = "none" | "first" | "last";
export type ExportPreset = "rich" | "plain";

export interface FieldOptions {
  mode?: FieldMode;
  variables?: Record<string, string | null>;
}

export interface HtmlOptions {
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
  trackedChangeMode?: "accept" | "reject";
  imagePlaceholder?: string;
  emitDocumentProperties?: boolean;
  emitCustomProperties?: boolean;
}

export type ExportRequest =
  | { format: "html"; preset?: ExportPreset; fields?: FieldOptions; html?: HtmlOptions }
  | { format: "markdown"; preset?: ExportPreset; fields?: FieldOptions; markdown?: MarkdownOptions }
  | { format: "text"; fields?: FieldOptions; text?: TextOptions };

export interface ResolveRequest { fields?: FieldOptions; }
export interface DocumentInfo { hasTrackedChanges: boolean; }

export interface Docxport {
  inspect(input: Uint8Array | ArrayBuffer): Promise<DocumentInfo>;
  export(input: Uint8Array | ArrayBuffer, request: ExportRequest): Promise<string>;
  resolveDocx(input: Uint8Array | ArrayBuffer, request?: ResolveRequest): Promise<Uint8Array>;
}

export function createDocxport(options?: DocxportInitOptions): Promise<Docxport>;
