import './csp-shims.js'
import { PptxViewer, RECOMMENDED_ZIP_LIMITS } from '@aiden0z/pptx-renderer'
import { MAX_DOCUMENT_BYTES, parseDocument } from './document-message.js'

const $ = (id) => document.getElementById(id)
const ui = Object.fromEntries(
  ['message', 'viewport', 'slides', 'prev', 'next', 'out', 'fit', 'in', 'percent', 'count', 'link-bar', 'link', 'close-link'].map((id) => [id, $(id)])
)
const MIN_ZOOM = 25
const MAX_ZOOM = 400
let version = -Infinity
let generation = 0
let viewer = null
let zoom = 100
let current = 0
const clamp = (value) => Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value)))

function controls() {
  const total = viewer?.slideCount ?? 0
  ui.prev.disabled = !viewer || current <= 0
  ui.next.disabled = !viewer || current >= total - 1
  ui.out.disabled = !viewer || zoom <= MIN_ZOOM
  ui.in.disabled = !viewer || zoom >= MAX_ZOOM
  ui.fit.disabled = !viewer
  ui.fit.setAttribute('aria-pressed', String(zoom === 100))
  ui.percent.textContent = `${zoom}%`
  ui.count.textContent = viewer ? `Slide ${current + 1} / ${total}` : ''
}

function showMessage(text, error = false) {
  ui.message.hidden = false
  ui.message.setAttribute('role', error ? 'alert' : 'status')
  ui.message.textContent = text
}

async function setZoom(next, anchor = null) {
  if (!viewer) return
  next = clamp(next)
  if (next === zoom) return
  const rect = ui.viewport.getBoundingClientRect()
  const point = anchor ?? { x: rect.width / 2, y: 0 }
  const x = ui.viewport.scrollLeft + point.x
  const y = ui.viewport.scrollTop + point.y
  const ratio = next / zoom
  zoom = next
  controls()
  await viewer.setZoom(zoom)
  ui.viewport.scrollLeft = x * ratio - point.x
  ui.viewport.scrollTop = y * ratio - point.y
}

async function go(index) {
  if (!viewer) return
  current = Math.min(viewer.slideCount - 1, Math.max(0, index))
  controls()
  await viewer.goToSlide(current, { behavior: 'instant', block: 'start' })
}

async function receive(doc) {
  if (doc.version <= version) return
  version = doc.version
  const token = ++generation
  const keep = viewer ? current : 0
  viewer?.destroy()
  viewer = null
  ui.slides.replaceChildren()
  ui['link-bar'].hidden = true
  controls()
  showMessage('Opening presentation…')
  if (doc.bytes.byteLength > MAX_DOCUMENT_BYTES) {
    showMessage('This presentation exceeds the 64 MiB preview limit.', true)
    return
  }
  try {
    const opened = await PptxViewer.open(doc.bytes, ui.slides, {
      renderMode: 'list',
      fitMode: 'contain',
      zoomPercent: zoom,
      scrollContainer: ui.viewport,
      zipLimits: RECOMMENDED_ZIP_LIMITS,
      lazyMedia: true,
      lazySlides: true,
      externalMedia: 'block',
      pdfjs: false,
      listOptions: { windowed: true, initialSlides: 4, batchSize: 4 },
      onSlideChange: (index) => {
        if (token !== generation) return
        current = index
        controls()
      }
    })
    if (token !== generation) {
      opened.destroy()
      return
    }
    viewer = opened
    if (viewer.slideCount === 0) throw new Error('This presentation has no slides.')
    ui.message.hidden = true
    current = Math.min(keep, viewer.slideCount - 1)
    controls()
    if (current > 0) await go(current)
  } catch (error) {
    if (token !== generation) return
    viewer?.destroy()
    viewer = null
    ui.slides.replaceChildren()
    controls()
    showMessage(
      /zip limit/i.test(error?.message ?? '')
        ? 'This presentation is too large to preview.'
        : 'Unable to preview this presentation. It may be damaged, password-protected or not a .pptx file.',
      true
    )
  }
}

// API 1.0 cannot open URLs, and following a link would navigate the editor frame away.
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

ui.slides.addEventListener('click', activateLink, true)
ui.slides.addEventListener('auxclick', activateLink, true)
ui.prev.addEventListener('click', () => go(current - 1))
ui.next.addEventListener('click', () => go(current + 1))
ui.out.addEventListener('click', () => setZoom(zoom / 1.25))
ui.in.addEventListener('click', () => setZoom(zoom * 1.25))
ui.fit.addEventListener('click', () => setZoom(100))
ui['close-link'].addEventListener('click', () => {
  ui['link-bar'].hidden = true
  ui.viewport.focus()
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
ui.viewport.addEventListener('wheel', (event) => {
  if (!viewer || !event.ctrlKey) return
  event.preventDefault()
  const rect = ui.viewport.getBoundingClientRect()
  const delta = event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? ui.viewport.clientHeight : 1)
  void setZoom(zoom * Math.exp(-Math.max(-100, Math.min(100, delta)) * 0.01), {
    x: event.clientX - rect.left,
    y: event.clientY - rect.top
  })
}, { passive: false })
controls()
showMessage('Opening presentation…')
window.orca.document.onDidChange((doc) => {
  const message = parseDocument(doc)
  if (message) void receive(message)
})
