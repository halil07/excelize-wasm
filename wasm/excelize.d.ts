// Type definitions for @halil07/excelize
// Project: https://github.com/halil07/excelize

export interface Options {
  password?: string;
  rawCellValue?: boolean;
  unzipSizeLimit?: number;
  unzipXMLSizeLimit?: number;
}

export interface Font {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  family?: string;
  size?: number;
  strike?: boolean;
  color?: string;
  colorIndexed?: number;
}

export interface Fill {
  type?: string;
  pattern?: number;
  color?: string[];
  shading?: number;
}

export interface Alignment {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  textRotation?: number;
  indent?: number;
  relativeIndent?: number;
  justifyLastLine?: boolean;
  shrinkToFit?: boolean;
  readingOrder?: number;
}

export interface Border {
  type?: string;
  color?: string;
  style?: number;
}

export interface Style {
  font?: Font;
  fill?: Fill;
  alignment?: Alignment;
  border?: Border[];
  numFmt?: number;
  decimalPlaces?: number;
  customNumFmt?: string;
  lang?: string;
  negred?: boolean;
  protection?: boolean;
}

export interface ConditionalFormatOptions {
  type?: string;
  aboveAverage?: boolean;
  percent?: boolean;
  format?: number;
  criteria?: string;
  value?: string;
  minType?: string;
  midType?: string;
  maxType?: string;
  minValue?: string;
  midValue?: string;
  maxValue?: string;
  minColor?: string;
  midColor?: string;
  maxColor?: string;
  barColor?: string;
  barBorderColor?: string;
  barDirection?: string;
  barOnly?: boolean;
  barSolid?: boolean;
  iconStyle?: string;
  reverseIcons?: boolean;
  iconsOnly?: boolean;
  stopIfTrue?: boolean;
}

export interface Table {
  range?: string;
  name?: string;
  styleName?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

export interface DataValidation {
  type?: string;
  operator?: string;
  formula1?: string;
  formula2?: string;
  sqref?: string;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  error?: string;
  errorTitle?: string;
  inputTitle?: string;
  inputBody?: string;
}

export interface GraphicOptions {
  altText?: string;
  printObject?: boolean;
  locked?: boolean;
  lockAspectRatio?: boolean;
  autoSize?: boolean;
  offsetX?: number;
  offsetY?: number;
  scaleX?: number;
  scaleY?: number;
  hyperlink?: string;
  hyperlinkType?: string;
  positioning?: string;
}

export interface ChartSeries {
  name?: string;
  categories?: string;
  values?: string;
}

export interface ChartAxis {
  title?: RichTextRun[];
  maximum?: number;
  minimum?: number;
  majorGridLines?: boolean;
  minorGridLines?: boolean;
  reverseOrder?: boolean;
  secondary?: boolean;
}

export interface ChartLegend {
  position?: string;
  showLegendKey?: boolean;
}

export interface ChartPlotArea {
  showBubbleSize?: boolean;
  showCatName?: boolean;
  showLeaderLines?: boolean;
  showPercent?: boolean;
  showSerName?: boolean;
  showVal?: boolean;
}

export interface RichTextRun {
  text?: string;
  font?: Font;
}

export interface Chart {
  type?: number;
  series?: ChartSeries[];
  format?: object;
  legend?: ChartLegend;
  title?: RichTextRun[];
  xAxis?: ChartAxis;
  yAxis?: ChartAxis;
  plotArea?: ChartPlotArea;
  showBlanksAs?: string;
  dimension?: object;
}

export interface PageLayoutOptions {
  size?: number;
  orientation?: string;
  firstPageNumber?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  pageOrder?: string;
  blackAndWhite?: boolean;
  cellComments?: string;
  copies?: number;
  draft?: boolean;
  horizontalDpi?: number;
  verticalDpi?: number;
}

export interface ViewOptions {
  defaultGridColor?: boolean;
  rightToLeft?: boolean;
  showFormulas?: boolean;
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showRuler?: boolean;
  showZeros?: boolean;
  topLeftCell?: string;
  view?: string;
  windowProtection?: boolean;
  zoomScale?: number;
}

export interface SheetPropsOptions {
  codeName?: string;
  enableFormatConditionsCalculation?: boolean;
  published?: boolean;
  autoPageBreaks?: boolean;
  fitToPage?: boolean;
  tabColor?: string;
  outlineProperties?: object;
  outlineSymbols?: boolean;
}

export interface AppProperties {
  application?: string;
  appVersion?: string;
  company?: string;
  docSecurity?: number;
  hyperlinksChanged?: boolean;
  linksUpToDate?: boolean;
  scaleCrop?: boolean;
  shareDoc?: boolean;
}

export interface DocProperties {
  category?: string;
  contentStatus?: string;
  created?: string;
  creator?: string;
  description?: string;
  identifier?: string;
  keywords?: string;
  lastModifiedBy?: string;
  modified?: string;
  revision?: string;
  subject?: string;
  title?: string;
  version?: string;
  language?: string;
}

export interface CalcPropsOptions {
  calcId?: string;
  calcMode?: string;
  fullCalcOnLoad?: boolean;
  refMode?: string;
  iterate?: boolean;
  iterateCount?: number;
  iterateDelta?: number;
}

export interface FormulaOpts {
  format?: string;
  shared?: boolean;
}

export interface HyperlinkOpts {
  display?: string;
  tooltip?: string;
  location?: string;
}

export interface HyperlinkResult {
  hasLink: boolean;
  link: string;
}

export interface MergeCell {
  upperLeft?: string;
  bottomRight?: string;
}

export interface TableResult {
  name?: string;
  range?: string;
  styleName?: string;
}

export interface PictureResult {
  extension?: string;
  file?: Uint8Array;
  format?: GraphicOptions;
}

export interface OkResult<T> {
  ok: true;
  data: T;
}

export interface ErrResult {
  ok: false;
  err: string;
}

export type Result<T> = OkResult<T> | ErrResult;

export interface RawBindings {
  newFile(opts?: Options | null): Result<number>;
  openFile(bytes: Uint8Array | ArrayBuffer, opts?: Options | null): Result<number>;
  close(id: number): Result<null>;
  save(id: number): Result<Uint8Array>;

  getSheetList(id: number): Result<string[]>;
  getSheetCount(id: number): Result<number>;
  defaultSheetName(id: number): Result<string>;

  newSheet(id: number, name: string): Result<number>;
  getSheetIndex(id: number, name: string): Result<number>;
  getSheetName(id: number, index: number): Result<string>;
  setSheetName(id: number, source: string, target: string): Result<null>;
  deleteSheet(id: number, name: string): Result<null>;
  setSheetVisible(id: number, name: string, visible: boolean, veryHidden?: boolean): Result<null>;
  getSheetVisible(id: number, name: string): Result<boolean>;
  getSheetMap(id: number): Result<Record<number, string>>;
  setActiveSheet(id: number, index: number): Result<null>;
  getActiveSheetIndex(id: number): Result<number>;
  setSheetBackgroundFromBytes(id: number, name: string, ext: string, bytes: Uint8Array | ArrayBuffer): Result<null>;
  setSheetDimension(id: number, name: string, range: string): Result<null>;
  getSheetDimension(id: number, name: string): Result<string>;

  setCellValue(id: number, sheet: string, cell: string, value: unknown): Result<null>;
  getCellValue(id: number, sheet: string, cell: string, opts?: Options | null): Result<unknown>;
  setCellFormula(id: number, sheet: string, cell: string, formula: string, opts?: FormulaOpts | null): Result<null>;
  getCellFormula(id: number, sheet: string, cell: string): Result<string>;
  setSheetRow(id: number, sheet: string, cell: string, values: unknown[]): Result<null>;
  setSheetCol(id: number, sheet: string, cell: string, values: unknown[]): Result<null>;
  getRows(id: number, sheet: string, opts?: Options | null): Result<string[][]>;
  getCols(id: number, sheet: string, opts?: Options | null): Result<string[][]>;

  setColWidth(id: number, sheet: string, c1: string, c2: string, width: number): Result<null>;
  getColWidth(id: number, sheet: string, col: string): Result<number>;
  setColVisible(id: number, sheet: string, cols: string, visible: boolean): Result<null>;
  getColVisible(id: number, sheet: string, col: string): Result<boolean>;
  setRowHeight(id: number, sheet: string, row: number, height: number): Result<null>;
  getRowHeight(id: number, sheet: string, row: number): Result<number>;
  setRowVisible(id: number, sheet: string, row: number, visible: boolean): Result<null>;
  getRowVisible(id: number, sheet: string, row: number): Result<boolean>;

  setColOutlineLevel(id: number, sheet: string, col: string, level: number): Result<null>;
  getColOutlineLevel(id: number, sheet: string, col: string): Result<number>;
  getRowOutlineLevel(id: number, sheet: string, row: number): Result<number>;

  newStyle(id: number, style: Style): Result<number>;
  getCellStyle(id: number, sheet: string, cell: string): Result<number>;
  setCellStyle(id: number, sheet: string, a: string, b: string, styleID: number): Result<null>;
  setConditionalFormat(id: number, sheet: string, range: string, opts: ConditionalFormatOptions[]): Result<null>;
  getConditionalFormats(id: number, sheet: string): Result<Record<string, ConditionalFormatOptions[]>>;

  mergeCell(id: number, sheet: string, a: string, b: string): Result<null>;
  unmergeCell(id: number, sheet: string, a: string, b: string): Result<null>;
  getMergeCells(id: number, sheet: string): Result<MergeCell[]>;

  addTable(id: number, sheet: string, table: Table): Result<null>;
  getTables(id: number, sheet: string): Result<TableResult[]>;
  deleteTable(id: number, name: string): Result<null>;

  addDataValidation(id: number, sheet: string, dv: DataValidation): Result<null>;

  addPictureFromBytes(id: number, sheet: string, cell: string, extension: string, bytes: Uint8Array | ArrayBuffer, graphicOptions?: GraphicOptions | null): Result<null>;
  getPictures(id: number, sheet: string, cell?: string): Result<PictureResult[]>;

  addChart(id: number, sheet: string, cell: string, chart: Chart): Result<null>;
  addChartSheet(id: number, sheet: string, chart: Chart): Result<null>;

  setCellHyperLink(id: number, sheet: string, cell: string, link: string, type: string, opts?: HyperlinkOpts | null): Result<null>;
  getCellHyperLink(id: number, sheet: string, cell: string): Result<HyperlinkResult>;

  setPageLayout(id: number, sheet: string, opts: PageLayoutOptions): Result<null>;
  getPageLayout(id: number, sheet: string): Result<PageLayoutOptions>;
  setSheetView(id: number, sheet: string, index: number, opts: ViewOptions): Result<null>;
  getSheetView(id: number, sheet: string, index: number): Result<ViewOptions>;
  setSheetProps(id: number, sheet: string, opts: SheetPropsOptions): Result<null>;
  getSheetProps(id: number, sheet: string): Result<SheetPropsOptions>;

  setAppProps(id: number, props: AppProperties): Result<null>;
  getAppProps(id: number): Result<AppProperties>;
  setDocProps(id: number, props: DocProperties): Result<null>;
  getDocProps(id: number): Result<DocProperties>;

  calcCellValue(id: number, sheet: string, cell: string, opts?: Options | null): Result<unknown>;
  setCalcProps(id: number, opts: CalcPropsOptions): Result<null>;
  getCalcProps(id: number): Result<CalcPropsOptions>;
}

export declare class File {
  constructor(id: number);
  close(): void;
  save(): Uint8Array;
  getSheetList(): string[];
  getSheetCount(): number;
  defaultSheetName(): string;

  newSheet(name: string): number;
  getSheetIndex(name: string): number;
  getSheetName(index: number): string;
  setSheetName(source: string, target: string): void;
  deleteSheet(name: string): void;
  setSheetVisible(name: string, visible: boolean, veryHidden?: boolean): void;
  getSheetVisible(name: string): boolean;
  getSheetMap(): Record<number, string>;
  setActiveSheet(index: number): void;
  getActiveSheetIndex(): number;
  setSheetBackgroundFromBytes(name: string, ext: string, bytes: Uint8Array | ArrayBuffer): void;
  setSheetDimension(name: string, range: string): void;
  getSheetDimension(name: string): string;

  setCellValue(sheet: string, cell: string, value: unknown): void;
  getCellValue(sheet: string, cell: string, opts?: Options): unknown;
  setCellFormula(sheet: string, cell: string, formula: string, opts?: FormulaOpts): void;
  getCellFormula(sheet: string, cell: string): string;
  setSheetRow(sheet: string, cell: string, values: unknown[]): void;
  setSheetCol(sheet: string, cell: string, values: unknown[]): void;
  getRows(sheet: string, opts?: Options): string[][];
  getCols(sheet: string, opts?: Options): string[][];

  setColWidth(sheet: string, c1: string, c2: string, width: number): void;
  getColWidth(sheet: string, col: string): number;
  setColVisible(sheet: string, cols: string, visible: boolean): void;
  getColVisible(sheet: string, col: string): boolean;
  setRowHeight(sheet: string, row: number, height: number): void;
  getRowHeight(sheet: string, row: number): number;
  setRowVisible(sheet: string, row: number, visible: boolean): void;
  getRowVisible(sheet: string, row: number): boolean;

  setColOutlineLevel(sheet: string, col: string, level: number): void;
  getColOutlineLevel(sheet: string, col: string): number;
  getRowOutlineLevel(sheet: string, row: number): number;

  newStyle(style: Style): number;
  getCellStyle(sheet: string, cell: string): number;
  setCellStyle(sheet: string, a: string, b: string, styleID: number): void;
  setConditionalFormat(sheet: string, range: string, opts: ConditionalFormatOptions[]): void;
  getConditionalFormats(sheet: string): Record<string, ConditionalFormatOptions[]>;

  mergeCell(sheet: string, a: string, b: string): void;
  unmergeCell(sheet: string, a: string, b: string): void;
  getMergeCells(sheet: string): MergeCell[];

  addTable(sheet: string, table: Table): void;
  getTables(sheet: string): TableResult[];
  deleteTable(name: string): void;

  addDataValidation(sheet: string, dv: DataValidation): void;

  addPictureFromBytes(sheet: string, cell: string, extension: string, bytes: Uint8Array | ArrayBuffer, graphicOptions?: GraphicOptions): void;
  getPictures(sheet: string, cell?: string): PictureResult[];

  addChart(sheet: string, cell: string, chart: Chart): void;
  addChartSheet(sheet: string, chart: Chart): void;

  setCellHyperLink(sheet: string, cell: string, link: string, type: string, opts?: HyperlinkOpts): void;
  getCellHyperLink(sheet: string, cell: string): HyperlinkResult;

  setPageLayout(sheet: string, opts: PageLayoutOptions): void;
  getPageLayout(sheet: string): PageLayoutOptions;
  setSheetView(sheet: string, index: number, opts: ViewOptions): void;
  getSheetView(sheet: string, index: number): ViewOptions;
  setSheetProps(sheet: string, opts: SheetPropsOptions): void;
  getSheetProps(sheet: string): SheetPropsOptions;

  setAppProps(props: AppProperties): void;
  getAppProps(): AppProperties;
  setDocProps(props: DocProperties): void;
  getDocProps(): DocProperties;

  calcCellValue(sheet: string, cell: string, opts?: Options): unknown;
  setCalcProps(opts: CalcPropsOptions): void;
  getCalcProps(): CalcPropsOptions;
}

export interface ExcelizeAPI {
  ready(): Promise<void>;
  excelizeInit(): Promise<ExcelizeAPI>;
  newFile(opts?: Options): File;
  openFile(bytes: Uint8Array | ArrayBuffer, opts?: Options): File;
  raw: RawBindings;
  File: typeof File;
}

export declare function excelizeInit(): Promise<ExcelizeAPI>;
export declare function ready(): Promise<void>;
export declare function newFile(opts?: Options): File;
export declare function openFile(bytes: Uint8Array | ArrayBuffer, opts?: Options): File;
export { ExcelizeAPI as Excelize };
export default ExcelizeAPI;
