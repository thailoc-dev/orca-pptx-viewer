import assert from 'node:assert/strict'
import { after, before, test } from 'node:test'
import { readFileSync } from 'node:fs'
import { asFormat, fixture, MAIN_TYPES } from './fixture.mjs'
import { launch, openEditor, postFile } from './harness.mjs'

let browser
let deck
before(async () => {
  browser = await launch()
  deck = await fixture()
})
after(async () => { await browser?.close() })

const opened = (frame) => frame.waitForFunction(() => document.getElementById('count').textContent !== '')

test('renders text, pictures, tables and charts in light and dark themes without CSP violations', async () => {
  for (const theme of ['light', 'dark']) {
    const { page, frame, errors } = await openEditor(browser, { theme })
    await postFile(page, deck)
    await opened(frame)
    assert.equal(await frame.locator('#count').textContent(), 'Slide 1 / 3')
    assert.equal(await frame.locator('#message').isHidden(), true)
    assert.match(await frame.locator('#slides').innerText(), /Quarterly review/)
    const picture = await frame.locator('#slides img').first().evaluate(async (image) => {
      await image.decode()
      return { data: image.src.startsWith('data:image/png'), width: image.naturalWidth }
    })
    assert.deepEqual(picture, { data: true, width: 8 })
    await frame.locator('#viewport').press('End')
    await frame.waitForFunction(() => document.querySelector('#slides canvas'))
    assert.match(await frame.locator('#slides').innerText(), /\$128,000/)
    assert.equal(await frame.locator('.footer .label').textContent(), 'Presentation · Read only')
    assert.deepEqual(errors, [])
    await page.close()
  }
})

test('navigates with buttons and keys, and zooms around fit', async () => {
  const { page, frame, errors } = await openEditor(browser, { width: 1000, height: 640 })
  await postFile(page, deck)
  await opened(frame)
  assert.equal(await frame.getByRole('button', { name: 'Previous slide' }).isDisabled(), true)
  await frame.getByRole('button', { name: 'Next slide' }).click()
  await frame.waitForFunction(() => document.getElementById('count').textContent === 'Slide 2 / 3')
  await frame.locator('#viewport').press('End')
  await frame.waitForFunction(() => document.getElementById('count').textContent === 'Slide 3 / 3')
  assert.equal(await frame.getByRole('button', { name: 'Next slide' }).isDisabled(), true)
  await frame.locator('#viewport').press('ArrowLeft')
  await frame.waitForFunction(() => document.getElementById('count').textContent === 'Slide 2 / 3')
  await frame.locator('#viewport').press('Home')
  await frame.waitForFunction(() => document.getElementById('count').textContent === 'Slide 1 / 3')

  const width = () => frame.locator('#viewport').evaluate((node) => node.scrollWidth)
  const fitted = await width()
  await frame.getByRole('button', { name: 'Zoom in', exact: true }).click()
  assert.equal(await frame.locator('#percent').textContent(), '125%')
  assert.ok(await width() > fitted)
  await frame.getByRole('button', { name: 'Fit to width', exact: true }).click()
  assert.equal(await frame.locator('#percent').textContent(), '100%')
  assert.equal(await width(), fitted)
  assert.deepEqual(errors, [])
  await page.close()
})

test('shows web links as copyable addresses instead of navigating the frame', async () => {
  const { page, frame, errors } = await openEditor(browser)
  await postFile(page, deck)
  await opened(frame)
  await frame.getByText('Visit the docs').click()
  assert.equal(await frame.locator('#link-bar').isVisible(), true)
  assert.equal(await frame.locator('#link').inputValue(), 'https://example.com/docs')
  assert.equal(await frame.locator('#count').textContent(), 'Slide 1 / 3')
  await frame.getByRole('button', { name: 'Close link address' }).click()
  assert.equal(await frame.locator('#link-bar').isHidden(), true)
  assert.deepEqual(errors, [])
  await page.close()
})

test('keeps the current slide when the file changes outside the editor', async () => {
  const { page, frame, errors } = await openEditor(browser)
  await postFile(page, deck)
  await opened(frame)
  await frame.getByRole('button', { name: 'Next slide' }).click()
  await frame.waitForFunction(() => document.getElementById('count').textContent === 'Slide 2 / 3')
  await postFile(page, deck)
  await frame.waitForFunction(() => document.getElementById('message').hidden)
  await frame.waitForFunction(() => document.getElementById('count').textContent === 'Slide 2 / 3')
  assert.deepEqual(errors, [])
  await page.close()
})

test('opens every format the manifest lists', async () => {
  const manifest = JSON.parse(readFileSync(new URL('../../orca-plugin.json', import.meta.url), 'utf8'))
  const extensions = manifest.contributes.customEditors[0].extensions
  assert.deepEqual(extensions, Object.keys(MAIN_TYPES))
  for (const extension of extensions) {
    const { page, frame, errors } = await openEditor(browser)
    await postFile(page, await asFormat(deck, extension))
    await opened(frame)
    assert.equal(await frame.locator('#count').textContent(), 'Slide 1 / 3', extension)
    assert.match(await frame.locator('#slides').innerText(), /Quarterly review/, extension)
    assert.deepEqual(errors, [], extension)
    await page.close()
  }
})

test('reports a damaged file instead of rendering', async () => {
  const { page, frame } = await openEditor(browser)
  await postFile(page, Buffer.from('this is not a presentation'))
  await frame.waitForFunction(() => document.getElementById('message').getAttribute('role') === 'alert')
  assert.match(await frame.locator('#message').textContent(), /Unable to preview this presentation/)
  assert.equal(await frame.locator('#count').textContent(), '')
  assert.equal(await frame.getByRole('button', { name: 'Zoom in', exact: true }).isDisabled(), true)
  await page.close()
})
