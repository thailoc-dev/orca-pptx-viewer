// Speaker notes. The renderer does not parse notes slides, so the notes part a slide links to is read
// from the package on demand, with its own ZIP reader opened once per file version.
import JSZip from 'jszip'

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NOTES_REL = /\/notesSlide$/
const MAX_NOTES_BYTES = 4 * 1024 * 1024

/** The package path of the notes part linked from a slide, or null when it has none. */
export function notesPathFor(slide) {
  if (!slide?.rels || !slide.slidePath) return null
  for (const rel of slide.rels.values()) {
    if (!NOTES_REL.test(rel.type) || /^external$/i.test(rel.targetMode ?? '')) continue
    return resolve(slide.slidePath, rel.target)
  }
  return null
}

function resolve(base, target) {
  const parts = target.startsWith('/') ? [] : base.split('/').slice(0, -1)
  for (const part of target.replace(/^\//, '').split('/')) {
    if (part === '..') parts.pop()
    else if (part && part !== '.') parts.push(part)
  }
  return parts.join('/')
}

const zips = new WeakMap()
function zipOf(bytes) {
  let zip = zips.get(bytes)
  if (!zip) {
    zip = JSZip.loadAsync(bytes)
    zips.set(bytes, zip)
  }
  return zip
}

/** The text of the notes body placeholder, one line per paragraph; '' when the slide has none. */
export async function readNotes(bytes, slide) {
  const path = notesPathFor(slide)
  if (!bytes || !path) return ''
  const entry = (await zipOf(bytes)).file(path)
  if (!entry) return ''
  // JSZip keeps the declared size from the central directory; refuse oversized parts before inflating.
  if ((entry._data?.uncompressedSize ?? 0) > MAX_NOTES_BYTES) return null
  const xml = await entry.async('string')
  if (xml.length > MAX_NOTES_BYTES) return null
  return notesText(xml)
}

export function notesText(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml')
  if (doc.getElementsByTagName('parsererror').length) return null
  const bodies = [...doc.getElementsByTagNameNS(P, 'sp')].filter((shape) => {
    const ph = shape.getElementsByTagNameNS(P, 'ph')[0]
    return ph && ph.getAttribute('type') === 'body'
  })
  const lines = []
  for (const body of bodies) {
    for (const paragraph of body.getElementsByTagNameNS(A, 'p')) {
      let line = ''
      for (const node of paragraph.getElementsByTagNameNS('*', '*')) {
        if (node.namespaceURI !== A) continue
        if (node.localName === 't') line += node.textContent
        else if (node.localName === 'br') line += '\n'
      }
      lines.push(line)
    }
  }
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim()
}
