// A stand-in for Orca's custom editor host, adapted from Orca's Word viewer tests: the same shell
// (CSP, design tokens, html class, and the document part of `window.orca`) in a sandboxed srcdoc
// frame, posting the file's bytes once the editor is ready. Headless Chromium only.
import { chromium } from 'playwright-core'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'

const HERE = import.meta.dirname
export const EDITOR_HTML = readFileSync(path.resolve(HERE, '..', '..', 'editor', 'index.html'), 'utf8')

// Same as PLUGIN_PANEL_CSP in Orca's src/shared/plugins/plugin-panel-shell.ts.
const CSP =
  "default-src 'none'; connect-src 'none'; script-src 'unsafe-inline'; " +
  "style-src 'unsafe-inline'; img-src data:; font-src data:; base-uri 'none'; form-action 'none'"

const TOKENS = {
  light: {
    background: '#fff', foreground: '#0a0a0a', muted: '#f5f5f5', 'muted-foreground': '#737373',
    accent: '#f5f5f5', border: '#e5e5e5', ring: '#a1a1a1'
  },
  dark: {
    background: '#0a0a0a', foreground: '#fafafa', muted: '#262626', 'muted-foreground': '#a1a1a1',
    accent: '#404040', border: 'rgb(255 255 255 / 0.07)', ring: '#737373'
  }
}

const ORCA_SHIM = `(function () {
  var listeners = []
  var post = function (message) { window.parent.postMessage(Object.assign({ orca: 1 }, message), '*') }
  document.addEventListener('securitypolicyviolation', function (e) {
    console.error('CSP violation: ' + e.violatedDirective + ' ' + e.blockedURI)
  })
  window.orca = Object.freeze({
    document: Object.freeze({
      onDidChange: function (fn) { listeners.push(fn); return { dispose: function () {} } }
    })
  })
  window.addEventListener('message', function (event) {
    var m = event.data
    if (event.source !== window.parent || !m || m.orca !== 1 || m.kind !== 'evt') return
    if (m.event === 'editor.document') listeners.forEach(function (fn) { fn(m.payload) })
  })
  post({ kind: 'evt', event: 'ready', payload: { protocol: 1 } })
})()`

function frameDocument(theme) {
  const tokens = Object.entries(TOKENS[theme]).map(([name, value]) => `--orca-${name}:${value};`).join('')
  return (
    `<!doctype html>\n<html class="${theme}">\n<head>\n<meta charset="utf-8">\n` +
    `<meta http-equiv="Content-Security-Policy" content="${CSP}">\n` +
    `<style id="orca-design-tokens">:root{${tokens}}</style>\n` +
    `<script>${ORCA_SHIM}</script>\n</head>\n${EDITOR_HTML}`
  )
}

const HOST_PAGE = `<!doctype html><html><head><style>html,body{margin:0;height:100%}iframe{border:0;width:100%;height:100%;display:block}</style></head>
<body><iframe id="editor" sandbox="allow-scripts"></iframe><script>
window.__readyCount = 0
window.addEventListener('message', (event) => {
  if (event.source !== document.getElementById('editor').contentWindow) return
  if (event.data && event.data.event === 'ready') window.__readyCount++
})
window.__version = 0
window.__postBytes = (base64) => {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  document.getElementById('editor').contentWindow.postMessage(
    { orca: 1, kind: 'evt', event: 'editor.document', payload: { version: ++window.__version, readOnly: false, fileType: '.pptx', content: { kind: 'bytes', bytes: bytes.buffer } } },
    '*',
    [bytes.buffer]
  )
}
</script></body></html>`

// Prefers the headless shell this playwright-core expects, then any other cached headless shell.
// Not Google Chrome: in its headless mode, clicks near the bottom of a sandboxed frame never arrive.
export async function launch() {
  try {
    return await chromium.launch({ headless: true })
  } catch (error) {
    const cache = process.env.PLAYWRIGHT_BROWSERS_PATH ?? path.join(os.homedir(), 'Library', 'Caches', 'ms-playwright')
    const shells = existsSync(cache) ? readdirSync(cache).filter((name) => name.startsWith('chromium_headless_shell-')).sort().reverse() : []
    for (const shell of shells) {
      const dir = path.join(cache, shell)
      for (const platform of readdirSync(dir)) {
        const executablePath = path.join(dir, platform, 'chrome-headless-shell')
        if (existsSync(executablePath)) return chromium.launch({ headless: true, executablePath })
      }
    }
    throw error
  }
}

export async function openEditor(browser, { theme = 'light', width = 1200, height = 800 } = {}) {
  const page = await browser.newPage({ viewport: { width, height }, deviceScaleFactor: 1 })
  const errors = []
  page.on('pageerror', (error) => errors.push(error.message))
  page.on('console', (message) => {
    if (message.type() === 'error') errors.push(message.text())
  })
  await page.setContent(HOST_PAGE)
  await page.evaluate((srcdoc) => (document.getElementById('editor').srcdoc = srcdoc), frameDocument(theme))
  await page.waitForFunction(() => window.__readyCount > 0)
  const frame = page.frames().find((candidate) => candidate !== page.mainFrame())
  return { page, frame, errors }
}

export async function postFile(page, bytes) {
  await page.evaluate((base64) => window.__postBytes(base64), Buffer.from(bytes).toString('base64'))
}
