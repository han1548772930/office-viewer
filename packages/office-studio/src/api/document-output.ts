import type { DocumentModel } from '../domain/model'
import { cloneModel } from '../domain/serialization'
import { documentDefaultStyle } from '../features/document/model/document-default-style'
import { documentUsesTraditionalPagination } from '../features/document/model/document-paragraph-pagination'
import { createRun } from '../features/document/model/document-model'
import { assertDocumentBodySerializable } from './document-body-codec'
import { bodyRecord, documentParagraphStyleToBody, documentRunStyleToBody } from './document-body-style'
import { documentSectionSetupToBody } from './document-body-section'
import { documentPublicBodyIndex } from './document-public-body'
import { writePublicDocumentStories } from './document-output-stories'
import { writePublicDocumentNotes } from './document-output-notes'
import { assertOfficeDocumentSnapshotCommentAnchors } from './document-snapshot-source'
import type { UniverDocumentData } from './document-snapshot-types'
import { readDocumentPluginResources, refreshDocumentLatexPluginWrapper, writeDocumentPluginResources } from './document-plugin-resources'
import { documentEffectiveZoomRatio } from '../features/document/services/document-zoom'

/** Detached v1.6.0 snapshot. It is built from the current command model, not
 * the unchanged input stream retained for Office/source preservation. */
export function documentToPublicSnapshot(model: DocumentModel, revision = 0): UniverDocumentData {
  if (!model.id?.trim()) throw new TypeError('文档快照需要稳定单元身份。')
  if (!Number.isSafeInteger(revision) || revision < 0) throw new RangeError('文档快照版本必须为非负安全整数。')
  assertOfficeDocumentSnapshotCommentAnchors(model)
  const defaults = documentDefaultStyle(model), source = cloneModel(bodyRecord(bodyRecord(model.metadata?.univerDocumentSource).snapshot))
  const sourceRevision = source.rev === undefined ? 0 : source.rev
  if (typeof sourceRevision !== 'number' || !Number.isSafeInteger(sourceRevision) || sourceRevision < 0 || !Number.isSafeInteger(sourceRevision + revision)) {
    throw new RangeError('文档公开版本或累计编辑版本超出安全整数范围。')
  }
  const data = documentPublicBodyIndex(model).data, sections = data.body.sectionBreaks ?? []
  const originalStyle = bodyRecord(source.documentStyle ?? model.metadata?.documentStyle)
  const modern = model.metadata?.documentFlavor === 2
  const style = documentSectionSetupToBody(modern ? { ...model.sectionSetup, ...model.page,
    orientation: model.metadata.documentPageOrientation === 'landscape' ? 'landscape' : 'portrait' } : { ...model.page, ...model.sectionSetup }, originalStyle)
  style.textStyle = documentRunStyleToBody(createRun('', { ...defaults.text, metadata: { univerTextStyle: bodyRecord(originalStyle.textStyle) } }))
  style.defaultParagraphStyle = documentParagraphStyleToBody(defaults.paragraph, originalStyle.defaultParagraphStyle)
  const flavor = model.metadata?.documentFlavor ?? originalStyle.documentFlavor ?? 1
  style.documentFlavor = flavor === 0 && documentUsesTraditionalPagination(model) ? 1 : flavor
  const stories = writePublicDocumentStories(model, source, sections), notes = writePublicDocumentNotes(model, source)
  const firstSection = sections[0] ?? {}
  for (const key of ['defaultHeaderId', 'defaultFooterId', 'firstPageHeaderId', 'firstPageFooterId', 'evenPageHeaderId', 'evenPageFooterId',
    'useFirstPageHeaderFooter', 'evenAndOddHeaders']) {
    if (firstSection[key] === undefined) delete style[key]
    else style[key] = firstSection[key]
  }
  const metadata = cloneModel(model.metadata ?? {})
  for (const key of ['univerDocumentSource', 'univerBodySource', 'univerBodyFirstSectionSource', 'documentEditingBlockHost',
    'documentEditingBlockSourceId', 'documentEditingSectionId', 'documentEditingSectionIds']) delete metadata[key]
  const snapshot: UniverDocumentData = { ...source, ...data, id: model.id, title: model.name, rev: sourceRevision + revision, documentStyle: style,
    settings: { ...bodyRecord(source.settings), zoomRatio: documentEffectiveZoomRatio(model) },
    styles: model.metadata?.documentStyles ? cloneModel(bodyRecord(model.metadata.documentStyles)) as UniverDocumentData['styles'] : undefined,
    tableSource: data.tableSource, drawings: data.drawings, drawingsOrder: data.drawingsOrder, lists: data.lists,
    callouts: undefined, codes: undefined, latexFormulas: undefined,
    resources: writeDocumentPluginResources(source, [data, ...Object.values(stories.headers), ...Object.values(stories.footers),
      ...Object.values(notes.notes), ...Object.values(notes.footnotes), ...Object.values(notes.endnotes)]),
    headers: Object.keys(stories.headers).length ? stories.headers : undefined, footers: Object.keys(stories.footers).length ? stories.footers : undefined,
    notes: Object.keys(notes.notes).length ? notes.notes : undefined,
    footnotes: Object.keys(notes.footnotes).length ? notes.footnotes : undefined, endnotes: Object.keys(notes.endnotes).length ? notes.endnotes : undefined,
    noteSettings: Object.keys(notes.noteSettings).length ? notes.noteSettings : undefined,
    customXmlParts: model.customXmlParts === undefined ? undefined : cloneModel(model.customXmlParts),
    officeStudio: { version: 1, metadata, ...(model.permissionRules === undefined ? {} : { permissionRules: cloneModel(model.permissionRules) }),
      ...(model.comments === undefined ? {} : { comments: cloneModel(model.comments) }) } }
  delete snapshot.text
  // Legacy input aliases and deleted modeled resources must not supersede
  // the current title/body when this public snapshot is reopened.
  delete snapshot.name
  const liveLatex = readDocumentPluginResources(snapshot.resources).latexFormulas ?? {}
  for (const entry of [snapshot, ...[snapshot.headers, snapshot.footers, snapshot.notes, snapshot.footnotes, snapshot.endnotes]
    .flatMap(group => Object.values(group ?? {}))]) {
    const value = bodyRecord(entry)
    delete value.latexFormulas
    if (value !== snapshot) refreshDocumentLatexPluginWrapper(value, liveLatex)
  }
  assertDocumentBodySerializable(snapshot)
  return snapshot
}
