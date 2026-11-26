import type { OfficeStudio } from './office-studio.ts'
import type { FPivotTable, FSheetPivotChart } from './facade.ts'
import { cloneModel } from '../domain/serialization.ts'
import type { TimelineLevel } from '../domain/model.ts'
import type { TimelinePublicLevel } from '../domain/timeline-filters.ts'
import { normalizeTimelineLevel, toTimelinePublicLevel } from '../features/slicer/timeline-date-filter.ts'
import { normalizeSheetSlicerAnchor, normalizeSheetSlicerInfo, normalizeSheetSlicerPlacement, normalizeSheetTimelineInfo, resolveSheetSlicerSource,
  type ISheetSlicerInfo, type ISheetTimelineInfo, type SheetSlicerAnchorSpec, type SheetSlicerPlacement, type SheetSlicerSourceSpec } from '../features/slicer/slicer-public-info.ts'
import type { SheetSlicerSettings } from '../features/slicer/slicer-settings.ts'
import type { SheetTimelineSettings } from '../features/slicer/timeline-settings.ts'

type CommonBuilderSettings = Partial<SheetSlicerSettings & SheetTimelineSettings>
type PivotReference = { unitId: string; subUnitId: string; pivotTableId: string }
type PivotChartReference = { unitId: string; subUnitId: string; pivotChartId: string }
type SourceSetter = 'setTableReference' | 'setPivotTable' | 'setPivotTableReference' | 'setPivotChart' | 'setPivotChartReference'
type TimelineSourceSetter = Exclude<SourceSetter, 'setTableReference'>
export type FLinkedSheetSlicerBuilder = {
  [K in keyof Omit<FSheetSlicerBuilder, SourceSetter>]: FSheetSlicerBuilder[K] extends (...args: infer A) => infer R
    ? R extends FSheetSlicerBuilder ? (...args: A) => FLinkedSheetSlicerBuilder : FSheetSlicerBuilder[K] : FSheetSlicerBuilder[K]
}
export type FBoundSheetSlicerBuilder = FLinkedSheetSlicerBuilder
export type FLinkedSheetTimelineBuilder = {
  [K in keyof Omit<FSheetTimelineBuilder, TimelineSourceSetter>]: FSheetTimelineBuilder[K] extends (...args: infer A) => infer R
    ? R extends FSheetTimelineBuilder ? (...args: A) => FLinkedSheetTimelineBuilder : FSheetTimelineBuilder[K] : FSheetTimelineBuilder[K]
}
export type FBoundSheetTimelineBuilder = FLinkedSheetTimelineBuilder

/** Detached configuration only. A source lock persists at runtime after every fluent call. */
abstract class SheetControlBuilder {
  protected source?: SheetSlicerSourceSpec
  protected settings: CommonBuilderSettings = {}
  protected placement: SheetSlicerPlacement = {}
  constructor(protected readonly api: OfficeStudio, protected readonly unitId: string, protected readonly fieldId: string, private readonly locked = false) {
    if (typeof fieldId !== 'string' || !fieldId.trim()) throw new TypeError('切片器 Builder 需要稳定来源字段标识。')
  }
  protected setSource(reference: { subUnitId: string; tableId: string } | PivotReference | PivotChartReference): this {
    if (this.locked) throw new Error('从活动控制取得的 Builder 来源已经锁定。')
    const workbook = this.api.getWorkbookModel(this.unitId)
    if (!workbook) throw new Error('切片器来源工作簿不存在。')
    this.source = resolveSheetSlicerSource(workbook, this.fieldId, reference)
    return this
  }
  setPivotTable(pivot: FPivotTable): this {
    return this.setPivotTableReference({ unitId: pivot.getUnitId(), subUnitId: pivot.getSubUnitId(), pivotTableId: pivot.getPivotTableId() })
  }
  setPivotTableReference(reference: PivotReference): this { return this.setSource(reference) }
  setPivotChart(chart: FSheetPivotChart): this {
    return this.setPivotChartReference({ unitId: chart.getUnitId(), subUnitId: chart.getSubUnitId(), pivotChartId: chart.getId() })
  }
  setPivotChartReference(reference: PivotChartReference): this { return this.setSource(reference) }
  setName(name: string): this { this.settings.name = name; return this }
  setCaption(caption: string): this { this.settings.caption = caption; return this }
  setStyle(style: string): this { this.settings.style = style; return this }
  clearStyle(): this { delete this.settings.style; return this }
  setPosition(anchor: SheetSlicerAnchorSpec): this {
    this.placement.anchor = normalizeSheetSlicerAnchor(anchor); delete this.placement.position; return this
  }
  setAbsolutePosition(x: number, y: number): this {
    this.placement.position = normalizeSheetSlicerPlacement({ position: { x, y } }).position; delete this.placement.anchor; return this
  }
  setSize(width: number, height: number): this { this.placement.size = normalizeSheetSlicerPlacement({ size: { width, height } }).size; return this }
}

export class FSheetSlicerBuilder extends SheetControlBuilder {
  constructor(api: OfficeStudio, unitId: string, fieldId: string, locked = false) {
    super(api, unitId, fieldId, locked); this.settings.rowHeight = 24
  }
  static fromInfo(api: OfficeStudio, unitId: string, info: ISheetSlicerInfo): FBoundSheetSlicerBuilder {
    const detached = normalizeSheetSlicerInfo(cloneModel(info)), fieldId = detached.dataSource.kind === 'table' ? detached.dataSource.columnId : detached.dataSource.sourceFieldId
    const builder = new FSheetSlicerBuilder(api, unitId, fieldId, true)
    builder.source = detached.dataSource; builder.settings = detached.settings
    const { dataSource: _source, settings: _settings, ...placement } = detached; builder.placement = placement
    return builder
  }
  setTableReference(reference: { subUnitId: string; tableId: string }): this { return this.setSource(reference) }
  setColumnCount(count: number): this { this.settings.columnCount = count; return this }
  setRowHeight(height: number): this { this.settings.rowHeight = height; return this }
  setShowCaption(visible: boolean): this { this.settings.showCaption = visible; return this }
  setLockedPosition(locked: boolean): this { this.settings.lockedPosition = locked; return this }
  setStartItem(index: number): this { this.settings.startItem = index; return this }
  build(): ISheetSlicerInfo {
    if (!this.source) throw new Error('插入切片器前需要配置来源。')
    return normalizeSheetSlicerInfo({ dataSource: cloneModel(this.source), settings: { ...cloneModel(this.settings), rowHeight: this.settings.rowHeight ?? 24 }, ...cloneModel(this.placement) })
  }
}

export class FSheetTimelineBuilder extends SheetControlBuilder {
  constructor(api: OfficeStudio, unitId: string, fieldId: string, locked = false) {
    super(api, unitId, fieldId, locked); this.settings.level = 2; this.settings.selectionLevel = 2
  }
  static fromInfo(api: OfficeStudio, unitId: string, info: ISheetTimelineInfo): FBoundSheetTimelineBuilder {
    const detached = normalizeSheetTimelineInfo(cloneModel(info)), builder = new FSheetTimelineBuilder(api, unitId, detached.dataSource.sourceFieldId, true)
    builder.source = detached.dataSource; builder.settings = detached.settings
    const { dataSource: _source, settings: _settings, ...placement } = detached; builder.placement = placement
    return builder
  }
  setLevel(level: TimelinePublicLevel | TimelineLevel): this { this.settings.level = toTimelinePublicLevel(normalizeTimelineLevel(level)); return this }
  setSelectionLevel(level: TimelinePublicLevel | TimelineLevel): this { this.settings.selectionLevel = toTimelinePublicLevel(normalizeTimelineLevel(level)); return this }
  setScrollPosition(position: number): this { this.settings.scrollPosition = position; return this }
  setShowHeader(visible: boolean): this { this.settings.showHeader = visible; return this }
  setShowSelectionLabel(visible: boolean): this { this.settings.showSelectionLabel = visible; return this }
  setShowTimeLevel(visible: boolean): this { this.settings.showTimeLevel = visible; return this }
  setShowHorizontalScrollbar(visible: boolean): this { this.settings.showHorizontalScrollbar = visible; return this }
  build(): ISheetTimelineInfo {
    if (!this.source || this.source.kind !== 'pivot') throw new Error('时间线需要配置原始透视日期来源。')
    return normalizeSheetTimelineInfo({ dataSource: cloneModel(this.source), settings: { ...cloneModel(this.settings),
      level: this.settings.level ?? 2, selectionLevel: this.settings.selectionLevel ?? 2 }, ...cloneModel(this.placement) })
  }
}
