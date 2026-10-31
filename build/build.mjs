// Builds ../editor/index.html: one self-contained file, because the editor frame's CSP allows only
// inline scripts and styles, data: images and fonts, with no network, workers or eval. Also writes
// ../THIRD_PARTY_NOTICES.txt from the packages that end up in the bundle.
// Usage: npm ci --ignore-scripts && node build.mjs
import { build, transform } from 'esbuild'
import { mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs'
import path from 'node:path'
import vm from 'node:vm'

const HERE = import.meta.dirname
const PLUGIN = path.resolve(HERE, '..')
const OUT = path.join(PLUGIN, 'editor', 'index.html')
const NOTICES = path.join(PLUGIN, 'THIRD_PARTY_NOTICES.txt')

const result = await build({
  entryPoints: [path.join(HERE, 'src', 'editor.js')],
  bundle: true,
  write: false,
  format: 'iife',
  platform: 'browser',
  mainFields: ['module', 'main'],
  target: 'chrome120',
  minify: true,
  legalComments: 'none',
  metafile: true,
  logLevel: 'warning',
  absWorkingDir: HERE,
  define: { 'process.env.NODE_ENV': '"production"' },
  alias: {
    setimmediate: path.join(HERE, 'src', 'set-immediate.js'),
    'readable-stream': path.join(HERE, 'src', 'no-node-stream.js'),
    stream: path.join(HERE, 'src', 'no-node-stream.js')
  },
  // JSZip's browser field selects a prebundled file that carries the setImmediate polyfill with its
  // new Function fallback. Bundle its original entry so the CSP-safe alias above applies instead.
  plugins: [{
    name: 'jszip-source',
    setup(builder) {
      builder.onResolve({ filter: /^jszip$/ }, () => ({
        path: path.join(HERE, 'node_modules', 'jszip', 'lib', 'index.js'), namespace: 'jszip-source'
      }))
      builder.onLoad({ filter: /.*/, namespace: 'jszip-source' }, ({ path: file }) => ({
        contents: readFileSync(file, 'utf8'), loader: 'js', resolveDir: path.dirname(file)
      }))
    }
  }]
})
const css = (
  await transform(readFileSync(path.join(HERE, 'src', 'editor.css'), 'utf8'), { loader: 'css', minify: true })
).code.trim()

// Why: `<!--` then `<script` inside an inline script puts the HTML parser in a state where `</script>`
// no longer ends it. Every occurrence is inside a string or regex literal, where `\x3C` reads the same.
const js = result.outputFiles[0].text.replace(/<(!--|script)/gi, '\\x3C$1').trim()
new vm.Script(js)
if (/<\/script/i.test(js) || /<\/style/i.test(css)) {
  throw new Error('the bundle contains a closing script or style tag; it cannot be inlined')
}
if (/\beval\(|new Function\(/.test(js)) {
  throw new Error('the bundle contains eval or new Function; the frame CSP refuses both')
}

const html = `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Presentation Viewer</title>
<style>${css}</style>
</head>
<body>
<div class="app">
<div class="content">
<div class="viewport" id="viewport" tabindex="0" aria-label="Presentation preview"><div class="slides" id="slides"></div></div>
<div class="message" id="message" role="status"></div>
</div>
<div class="link-bar" id="link-bar" hidden><label for="link">Copy link address</label><input id="link" readonly aria-label="Link address"><button id="close-link" type="button" aria-label="Close link address">×</button></div>
<div class="footer">
<button id="prev" type="button" title="Previous slide" aria-label="Previous slide">‹</button>
<button id="next" type="button" title="Next slide" aria-label="Next slide">›</button>
<span class="separator"></span>
<button id="out" type="button" title="Zoom out" aria-label="Zoom out">−</button>
<button id="fit" type="button" title="Fit to width" aria-label="Fit to width">Fit</button>
<button id="in" type="button" title="Zoom in" aria-label="Zoom in">+</button>
<span id="percent" class="percent">100%</span>
<span class="label">Presentation · Read only</span>
<span id="count" class="slide-count" aria-live="polite"></span>
</div>
</div>
<script>${js}</script>
</body>
</html>
`
mkdirSync(path.dirname(OUT), { recursive: true })
writeFileSync(OUT, html)

// Notices: every npm package with code in the bundle, with the licence text it ships.
const packages = new Map()
for (const rawInput of Object.keys(result.metafile.inputs)) {
  const input = rawInput.replace(/^jszip-source:/, '')
  const at = input.lastIndexOf('node_modules/')
  if (at === -1) continue
  const name = /^(?:@[^/]+\/)?[^/]+/.exec(input.slice(at + 'node_modules/'.length))[0]
  const dir = path.resolve(HERE, input.slice(0, at), 'node_modules', name)
  packages.set(dir, { ...JSON.parse(readFileSync(path.join(dir, 'package.json'), 'utf8')), dir })
}
const licenceText = (pkg) => {
  const file = readdirSync(pkg.dir).find((entry) => /^(licen[cs]e|copying)(\.|$)/i.test(entry))
  if (!file) throw new Error(`${pkg.name} ships no licence file`)
  const text = readFileSync(path.join(pkg.dir, file), 'utf8').replace(/\r\n/g, '\n').trim()
  // Apache-2.0 asks redistributions to carry the package's NOTICE file along with its licence.
  const notice = readdirSync(pkg.dir).find((entry) => /^notice(\.|$)/i.test(entry))
  return notice ? `${text}\n\n${readFileSync(path.join(pkg.dir, notice), 'utf8').replace(/\r\n/g, '\n').trim()}` : text
}
const sections = [...packages.values()]
  .sort((a, b) => a.name.localeCompare(b.name))
  .map((pkg) => `${pkg.name}@${pkg.version} (${pkg.license})\n${pkg.homepage ?? ''}\n\n${licenceText(pkg)}`)
writeFileSync(
  NOTICES,
  `Presentation Viewer's editor/index.html bundles the following software.\n\n` +
    `${sections.join(`\n\n${'='.repeat(78)}\n\n`)}\n`
)

console.log(`editor/index.html: ${Buffer.byteLength(html)} bytes, ${packages.size} packages`)
