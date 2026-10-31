// A synthetic three-slide deck: title text, a picture, a table, a shape, a chart and a hyperlink.
import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'

const MAIN_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'

// The slide show, template and macro-enabled formats differ from .pptx only in the content type of
// ppt/presentation.xml (and, when macro-enabled, a VBA part the viewer never reads).
export const MAIN_TYPES = {
  '.pptx': MAIN_TYPE,
  '.ppsx': 'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml',
  '.potx': 'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
  '.pptm': 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  '.ppsm': 'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
  '.potm': 'application/vnd.ms-powerpoint.template.macroEnabled.main+xml'
}

/** The same deck saved as another PowerPoint format, such as `.ppsx`. */
export async function asFormat(deck, extension) {
  const zip = await JSZip.loadAsync(deck)
  const types = await zip.file('[Content_Types].xml').async('string')
  if (!types.includes(MAIN_TYPE)) throw new Error('the fixture has no presentation part')
  zip.file('[Content_Types].xml', types.replace(MAIN_TYPE, MAIN_TYPES[extension]))
  return zip.generateAsync({ type: 'nodebuffer' })
}

// An 8×8 red PNG.
export const RED_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGO4o6GBFTEMLQkAe3tLAfuiUfAAAAAASUVORK5CYII='

export async function fixture() {
  const pptx = new PptxGenJS()
  pptx.layout = 'LAYOUT_16x9'

  const title = pptx.addSlide()
  title.background = { color: 'FBF5E6' }
  title.addText('Quarterly review', { x: 0.5, y: 1.2, w: 9, h: 1, fontSize: 40, bold: true, color: 'B8461F' })
  title.addText('Visit the docs', { x: 0.5, y: 2.6, w: 9, h: 0.6, fontSize: 18, hyperlink: { url: 'https://example.com/docs' } })
  title.addImage({ data: `image/png;base64,${RED_PNG}`, x: 7.5, y: 3.5, w: 1.5, h: 1.5 })

  const table = pptx.addSlide()
  table.addText('Revenue', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 28 })
  table.addTable(
    [
      [{ text: 'Region' }, { text: 'Q3' }],
      [{ text: 'North' }, { text: '$128,000' }],
      [{ text: 'South' }, { text: '$96,500' }]
    ],
    { x: 0.5, y: 1.2, w: 6, fontSize: 16, border: { type: 'solid', pt: 1, color: '888888' } }
  )
  table.addShape(pptx.ShapeType.roundRect, { x: 7, y: 1.2, w: 2.5, h: 1.5, fill: { color: '1F6F7A' }, rectRadius: 0.2 })

  const chart = pptx.addSlide()
  chart.addText('Growth', { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 28 })
  chart.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2', 'Q3'], values: [40, 65, 90] }], {
    x: 0.5, y: 1.2, w: 8, h: 4
  })

  return pptx.write({ outputType: 'nodebuffer' })
}
