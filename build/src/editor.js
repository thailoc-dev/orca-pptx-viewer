import './csp-shims.js'
import { PptxViewer, RECOMMENDED_ZIP_LIMITS, buildPresentation, parseZipLazyMedia } from '@aiden0z/pptx-renderer'
import { MAX_DOCUMENT_BYTES, parseDocument } from './document-message.js'
import { notesPathFor, readNotes } from './notes.js'

const $ = (id) => document.getElementById(id)
const ui = Object.fromEntries(
  [
    'message', 'viewport', 'slides', 'prev', 'next', 'out', 'fit', 'page', 'in', 'percent', 'count', 'status',
    'slide-input', 'slide-total', 'link-bar', 'link', 'close-link', 'thumbs', 'toggle-thumbs',
    'search-bar', 'search', 'search-count', 'search-prev', 'search-next', 'close-search', 'toggle-search',
    'notes', 'notes-title', 'notes-text', 'toggle-notes'
  ].map((id) => [id, $(id)])
)
const MIN_ZOOM = 25
const MAX_ZOOM = 400
const LIST_OPTIONS = { windowed: true, initialSlides: 4, batchSize: 4 }
const HIGHLIGHT = {
  borderColor: '#f59e0b', backgroundColor: 'rgba(245, 158, 11, 0.2)', borderWidth: 2, borderRadius: 3, padding: 2
}

let version = -Infinity
let generation = 0
let viewer = null
let bytes = null
let zoom = 100
// 'width' and 'page' follow the editor's size; 'custom' keeps the chosen percentage.
let fit = 'width'
let current = 0
const clamp = (value) => Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value)))

// ---------------------------------------------------------------------------
// Controls

function controls() {
  const total = viewer?.slideCount ?? 0
  ui.prev.disabled = !viewer || current <= 0
  ui.next.disabled = !viewer || current >= total - 1
  ui.out.disabled = !viewer || zoom <= MIN_ZOOM
  ui.in.disabled = !viewer || zoom >= MAX_ZOOM
  for (const id of ['fit', 'page', 'slide-input', 'toggle-search']) ui[id].disabled = !viewer
  ui.fit.setAttribute('aria-pressed', String(fit === 'width'))
  ui.page.setAttribute('aria-pressed', String(fit === 'page'))
  ui.percent.textContent = `${zoom}%`
  ui.count.textContent = viewer ? `Slide ${current + 1} / ${total}` : ''
  ui['slide-total'].textContent = viewer ? `/ ${total}` : ''
  if (document.activeElement !== ui['slide-input']) ui['slide-input'].value = viewer ? String(current + 1) : ''
  ui['slide-input'].style.width = `${Math.max(2, String(total).length) + 1.5}ch`
  markThumb()
  void showNotes()
}

function showMessage(text, error = false) {
  ui.message.hidden = false
  ui.message.setAttribute('role', error ? 'alert' : 'status')
  ui.message.textContent = text
}

function setStatus(text = '') {
  ui.status.textContent = text
  ui.status.hidden = !text
}

// ---------------------------------------------------------------------------
// Zoom. Wheel and pinch events arrive faster than the renderer re-lays out slides, so requests are
// coalesced: only the latest target is applied once the previous zoom has finished.

let zoomTarget = null
let zooming = false

function requestZoom(next, anchor = null, mode = 'custom') {
  if (!viewer) return
  fit = mode
  zoomTarget = { value: clamp(next), anchor }
  if (!zooming) void drainZoom()
  else controls()
}

async function drainZoom() {
  zooming = true
  try {
    while (zoomTarget && viewer) {
      const { value, anchor } = zoomTarget
      zoomTarget = null
      await applyZoom(value, anchor)
      await new Promise(requestAnimationFrame)
    }
  } finally {
    zooming = false
    controls()
  }
}

async function applyZoom(next, anchor) {
  if (next === zoom) return
  const rect = ui.viewport.getBoundingClientRect()
  const point = anchor ?? { x: rect.width / 2, y: 0 }
  const x = ui.viewport.scrollLeft + point.x
  const y = ui.viewport.scrollTop + point.y
  const ratio = next / zoom
  zoom = next
  controls()
  const target = viewer
  await target.setZoom(zoom)
  if (target !== viewer) return
  ui.viewport.scrollLeft = x * ratio - point.x
  ui.viewport.scrollTop = y * ratio - point.y
  if (fit === 'page') await target.goToSlide(current, { behavior: 'instant', block: 'center' })
  void rehighlight()
}

// At 100% the renderer fits a slide to the viewport's width; this is the zoom at which the whole
// slide, plus its 20px gap, fits the viewport's height too. Never larger than fitting the width.
function pageZoom() {
  if (!viewer) return 100
  const width = ui.slides.clientWidth || ui.viewport.clientWidth
  const height = (width * viewer.slideHeight) / viewer.slideWidth
  if (!height) return 100
  return Math.min(100, Math.floor((100 * (ui.viewport.clientHeight - 28)) / height))
}

function refit() {
  if (fit === 'page') requestZoom(pageZoom(), null, 'page')
}

// ---------------------------------------------------------------------------
// Navigation

async function go(index, focusThumb = false) {
  if (!viewer) return
  current = Math.min(viewer.slideCount - 1, Math.max(0, index))
  controls()
  if (focusThumb) thumbButtons[current]?.focus()
  await viewer.goToSlide(current, { behavior: 'instant', block: fit === 'page' ? 'center' : 'start' })
}

function commitSlideInput() {
  const value = Number.parseInt(ui['slide-input'].value, 10)
  if (viewer && Number.isFinite(value)) void go(value - 1)
  ui['slide-input'].value = viewer ? String(current + 1) : ''
}

// ---------------------------------------------------------------------------
// Opening. The new version is parsed while the previous one stays on screen, so a file saved from
// PowerPoint does not flash; a version that fails to parse leaves the previous one in place.

function teardown() {
  clearThumbs()
  search.results = []
  search.index = -1
  notesCache.clear()
  viewer?.destroy()
  viewer = null
  bytes = null
  ui.slides.replaceChildren()
}

async function receive(doc) {
  if (doc.version <= version) return
  version = doc.version
  const token = ++generation
  const reloading = viewer !== null
  ui['link-bar'].hidden = true
  if (doc.bytes.byteLength > MAX_DOCUMENT_BYTES) {
    teardown()
    controls()
    showMessage('This presentation exceeds the 64 MiB preview limit.', true)
    return
  }
  if (reloading) setStatus('Reloading…')
  else showMessage('Opening presentation…')
  try {
    const files = await parseZipLazyMedia(doc.bytes, RECOMMENDED_ZIP_LIMITS)
    if (token !== generation) return
    const presentation = buildPresentation(files, { lazySlides: true })
    if (presentation.slides.length === 0) throw new Error('This presentation has no slides.')
    const keep = current
    teardown()
    // The renderer reports slide 0 after its first render; ignore that until the kept slide is shown.
    let settled = false
    const created = new PptxViewer(ui.slides, {
      fitMode: 'contain',
      zoomPercent: zoom,
      scrollContainer: ui.viewport,
      externalMedia: 'block',
      pdfjs: false,
      onSlideChange: (index) => {
        if (created !== viewer || !settled || index === current) return
        current = index
        controls()
      }
    })
    viewer = created
    bytes = doc.bytes
    current = Math.min(keep, presentation.slides.length - 1)
    created.load(presentation)
    await created.renderList(LIST_OPTIONS)
    if (created !== viewer) return
    ui.message.hidden = true
    setStatus()
    ui['toggle-notes'].dataset.empty = String(!presentation.slides.some((slide) => notesPathFor(slide)))
    controls()
    buildThumbs()
    if (current > 0) await go(current)
    settled = true
    if (fit === 'page') refit()
    if (!ui['search-bar'].hidden) runSearch(current)
  } catch (error) {
    if (token !== generation) return
    if (reloading && viewer) {
      setStatus('The changed file could not be read; showing the previous version.')
      return
    }
    teardown()
    controls()
    setStatus()
    showMessage(
      /zip limit/i.test(error?.message ?? '')
        ? 'This presentation is too large to preview.'
        : /no slides/i.test(error?.message ?? '')
          ? 'This presentation has no slides.'
          : 'Unable to preview this presentation. It may be damaged, password-protected or not a .pptx file.',
      true
    )
  }
}

// ---------------------------------------------------------------------------
// Slide thumbnails. Previews mount only near the visible part of the sidebar and are disposed when
// they scroll away, so a long deck never holds more than a few dozen rendered previews.

let thumbButtons = []
let thumbObserver = null
const thumbHandles = new Map()

function clearThumbs() {
  thumbObserver?.disconnect()
  thumbObserver = null
  for (const handle of thumbHandles.values()) handle.dispose()
  thumbHandles.clear()
  thumbButtons = []
  ui.thumbs.replaceChildren()
}

function buildThumbs() {
  clearThumbs()
  if (!viewer || ui.thumbs.hidden) return
  const owner = viewer
  const width = Math.max(64, ui.thumbs.clientWidth - 26)
  const height = Math.round((width * owner.slideHeight) / owner.slideWidth)
  thumbObserver = new IntersectionObserver((entries) => {
    if (owner !== viewer) return
    for (const entry of entries) {
      const index = Number(entry.target.dataset.index)
      if (entry.isIntersecting && !thumbHandles.has(index)) {
        const handle = owner.renderThumbnailToContainer(index, entry.target.firstElementChild, { width })
        if (handle) thumbHandles.set(index, handle)
      } else if (!entry.isIntersecting && thumbHandles.has(index)) {
        thumbHandles.get(index).dispose()
        thumbHandles.delete(index)
      }
    }
  }, { root: ui.thumbs, rootMargin: '400px 0px' })
  const fragment = document.createDocumentFragment()
  owner.presentationData.slides.forEach((slide, index) => {
    const button = document.createElement('button')
    button.type = 'button'
    button.className = 'thumb'
    button.dataset.index = String(index)
    button.setAttribute('aria-label', `Slide ${index + 1}${slide.hidden ? ' (hidden)' : ''}`)
    const frame = document.createElement('span')
    frame.className = 'thumb-frame'
    frame.style.width = `${width}px`
    frame.style.height = `${height}px`
    const number = document.createElement('span')
    number.className = 'thumb-number'
    number.textContent = slide.hidden ? `${index + 1} · Hidden` : String(index + 1)
    if (slide.hidden) button.classList.add('hidden-slide')
    button.append(frame, number)
    fragment.append(button)
    thumbButtons.push(button)
  })
  ui.thumbs.append(fragment)
  for (const button of thumbButtons) thumbObserver.observe(button)
  markThumb(true)
}

function markThumb(force = false) {
  thumbButtons.forEach((button, index) => {
    const active = index === current
    if (active === (button.getAttribute('aria-current') === 'true') && !force) return
    if (active) {
      button.setAttribute('aria-current', 'true')
      button.tabIndex = 0
      button.scrollIntoView({ block: 'nearest' })
    } else {
      button.removeAttribute('aria-current')
      button.tabIndex = -1
    }
  })
}

function toggleThumbs(open = ui.thumbs.hidden) {
  ui.thumbs.hidden = !open
  ui['toggle-thumbs'].setAttribute('aria-pressed', String(open))
  if (open) buildThumbs()
  else clearThumbs()
}

// ---------------------------------------------------------------------------
// Find. Searches the parsed text of every slide, not only the mounted ones, and outlines the shape
// or table that contains each match.

const search = { results: [], index: -1, timer: 0 }

function runSearch(from = current) {
  clearTimeout(search.timer)
  viewer?.clearSearchHighlights()
  const query = ui.search.value
  search.results = viewer && query.trim() ? viewer.searchText(query) : []
  search.index = -1
  if (search.results.length) {
    const next = search.results.findIndex((result) => result.slideIndex >= from)
    void showResult(next === -1 ? 0 : next)
  } else {
    searchCount()
  }
}

function searchCount() {
  const { results, index } = search
  ui['search-count'].textContent = !ui.search.value.trim() ? '' : results.length ? `${index + 1} of ${results.length}` : 'No results'
  ui['search-prev'].disabled = ui['search-next'].disabled = results.length < 2
}

async function showResult(index) {
  if (!viewer || !search.results.length) return
  const count = search.results.length
  search.index = ((index % count) + count) % count
  searchCount()
  const result = search.results[search.index]
  const owner = viewer
  owner.clearSearchHighlights()
  current = result.slideIndex
  controls()
  await owner.highlightSearchResult(result, { ...HIGHLIGHT, scrollIntoView: { behavior: 'instant', block: 'center' } })
}

async function rehighlight() {
  const result = search.results[search.index]
  if (!viewer || !result || ui['search-bar'].hidden) return
  viewer.clearSearchHighlights()
  await viewer.highlightSearchResult(result, { ...HIGHLIGHT, scrollIntoView: false })
}

function openSearch() {
  if (!viewer) return
  ui['search-bar'].hidden = false
  ui['toggle-search'].setAttribute('aria-pressed', 'true')
  ui.search.focus()
  ui.search.select()
  if (ui.search.value.trim()) runSearch(current)
}

function closeSearch() {
  clearTimeout(search.timer)
  ui['search-bar'].hidden = true
  ui['toggle-search'].setAttribute('aria-pressed', 'false')
  viewer?.clearSearchHighlights()
  search.results = []
  search.index = -1
  ui.viewport.focus()
}

// ---------------------------------------------------------------------------
// Speaker notes, read from the notes part of the current slide when the panel is open.

const notesCache = new Map()
let notesRequest = 0

async function showNotes() {
  if (ui.notes.hidden) return
  const request = ++notesRequest
  if (!viewer) {
    ui['notes-title'].textContent = 'Notes'
    ui['notes-text'].textContent = ''
    return
  }
  const index = current
  ui['notes-title'].textContent = `Notes · Slide ${index + 1}`
  const slide = viewer.presentationData.slides[index]
  let text = notesCache.get(index)
  if (text === undefined) {
    const owner = viewer
    text = await readNotes(bytes, slide).catch(() => null)
    if (owner !== viewer) return
    notesCache.set(index, text)
  }
  if (request !== notesRequest) return
  ui['notes-text'].classList.toggle('empty', !text)
  ui['notes-text'].textContent = text === null ? 'The notes for this slide could not be read.' : text || 'No speaker notes for this slide.'
}

function toggleNotes(open = ui.notes.hidden) {
  ui.notes.hidden = !open
  ui['toggle-notes'].setAttribute('aria-pressed', String(open))
  void showNotes()
}

// ---------------------------------------------------------------------------
// Links. API 1.0 cannot open URLs, and following a link would navigate the editor frame away.
// Show the address so it can be copied into a browser instead.

function activateLink(event) {
  const anchor = event.target instanceof Element ? event.target.closest('a[href]') : null
  if (!anchor) return
  event.preventDefault()
  event.stopPropagation()
  const href = anchor.getAttribute('href') ?? ''
  if (/^https?:/i.test(href) && event.type === 'click') {
    ui['link-bar'].hidden = false
    ui.link.value = href
    ui.link.focus()
    ui.link.select()
  }
}

// ---------------------------------------------------------------------------
// Events

ui.slides.addEventListener('click', activateLink, true)
ui.slides.addEventListener('auxclick', activateLink, true)
ui.prev.addEventListener('click', () => go(current - 1))
ui.next.addEventListener('click', () => go(current + 1))
ui.out.addEventListener('click', () => requestZoom(zoom / 1.25))
ui.in.addEventListener('click', () => requestZoom(zoom * 1.25))
ui.fit.addEventListener('click', () => requestZoom(100, null, 'width'))
ui.page.addEventListener('click', () => requestZoom(pageZoom(), null, 'page'))
ui['close-link'].addEventListener('click', () => {
  ui['link-bar'].hidden = true
  ui.viewport.focus()
})
ui['slide-input'].addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault()
    commitSlideInput()
    ui.viewport.focus()
  } else if (event.key === 'Escape') {
    ui['slide-input'].value = String(current + 1)
    ui.viewport.focus()
  }
})
ui['slide-input'].addEventListener('change', commitSlideInput)
ui['slide-input'].addEventListener('focus', () => ui['slide-input'].select())
ui['toggle-thumbs'].addEventListener('click', () => toggleThumbs())
ui['toggle-notes'].addEventListener('click', () => toggleNotes())
ui['toggle-search'].addEventListener('click', () => (ui['search-bar'].hidden ? openSearch() : closeSearch()))
ui['close-search'].addEventListener('click', closeSearch)
ui['search-prev'].addEventListener('click', () => showResult(search.index - 1))
ui['search-next'].addEventListener('click', () => showResult(search.index + 1))
ui.search.addEventListener('input', () => {
  clearTimeout(search.timer)
  search.timer = setTimeout(() => runSearch(current), 120)
})
ui.search.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault()
    if (search.timer && search.index === -1) runSearch(current)
    else void showResult(search.index + (event.shiftKey ? -1 : 1))
  } else if (event.key === 'Escape') {
    event.preventDefault()
    closeSearch()
  }
})
ui.thumbs.addEventListener('click', (event) => {
  const button = event.target instanceof Element ? event.target.closest('.thumb') : null
  if (button) void go(Number(button.dataset.index))
})
ui.thumbs.addEventListener('keydown', (event) => {
  if (!viewer || event.metaKey || event.ctrlKey || event.altKey) return
  const step = { ArrowUp: -1, ArrowDown: 1, ArrowLeft: -1, ArrowRight: 1, PageUp: -1, PageDown: 1 }[event.key]
  if (step) {
    event.preventDefault()
    void go(current + step, true)
  } else if (event.key === 'Home' || event.key === 'End') {
    event.preventDefault()
    void go(event.key === 'Home' ? 0 : viewer.slideCount - 1, true)
  }
})
ui.viewport.addEventListener('keydown', (event) => {
  if (!viewer || event.metaKey || event.ctrlKey || event.altKey) return
  const step = { ArrowLeft: -1, ArrowRight: 1, PageUp: -1, PageDown: 1 }[event.key]
  if (step) {
    event.preventDefault()
    void go(current + step)
  } else if (event.key === 'Home' || event.key === 'End') {
    event.preventDefault()
    void go(event.key === 'Home' ? 0 : viewer.slideCount - 1)
  }
})
document.addEventListener('keydown', (event) => {
  const mod = event.metaKey || event.ctrlKey
  if (mod && !event.altKey && event.key.toLowerCase() === 'f') {
    event.preventDefault()
    openSearch()
  } else if (mod && !event.altKey && (event.key === '=' || event.key === '+')) {
    event.preventDefault()
    requestZoom(zoom * 1.25)
  } else if (mod && !event.altKey && event.key === '-') {
    event.preventDefault()
    requestZoom(zoom / 1.25)
  } else if (mod && !event.altKey && event.key === '0') {
    event.preventDefault()
    requestZoom(100, null, 'width')
  } else if (event.key === 'Escape' && !ui['link-bar'].hidden) {
    ui['link-bar'].hidden = true
    ui.viewport.focus()
  }
})
ui.viewport.addEventListener('wheel', (event) => {
  if (!viewer || !event.ctrlKey) return
  event.preventDefault()
  const rect = ui.viewport.getBoundingClientRect()
  const delta = event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? ui.viewport.clientHeight : 1)
  const base = zoomTarget?.value ?? zoom
  requestZoom(base * Math.exp(-Math.max(-100, Math.min(100, delta)) * 0.01), {
    x: event.clientX - rect.left,
    y: event.clientY - rect.top
  })
}, { passive: false })
let resizeFrame = 0
new ResizeObserver(() => {
  cancelAnimationFrame(resizeFrame)
  resizeFrame = requestAnimationFrame(refit)
}).observe(ui.viewport)

controls()
showMessage('Opening presentation…')
window.orca.document.onDidChange((doc) => {
  const message = parseDocument(doc)
  if (message) void receive(message)
})
