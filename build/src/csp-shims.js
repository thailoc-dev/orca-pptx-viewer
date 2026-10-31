// The editor frame's CSP loads images only from data: URLs, so the blob: URLs the renderer makes
// for slide pictures would never load. Blobs built from bytes remember them, and createObjectURL
// turns such a Blob into a data: URL synchronously, keeping the renderer's call sites unchanged.
const NativeBlob = globalThis.Blob
const nativeCreate = URL.createObjectURL.bind(URL)
const nativeRevoke = URL.revokeObjectURL.bind(URL)
const tracked = new WeakMap()
const encoder = new TextEncoder()

function bytesOfPart(part) {
  if (part instanceof ArrayBuffer) return new Uint8Array(part)
  if (ArrayBuffer.isView(part)) return new Uint8Array(part.buffer, part.byteOffset, part.byteLength)
  if (typeof part === 'string') return encoder.encode(part)
  return tracked.get(part) ?? null
}

class TrackedBlob extends NativeBlob {
  constructor(parts = [], options) {
    super(parts, options)
    const chunks = []
    for (const part of parts) {
      const bytes = bytesOfPart(part)
      if (!bytes) return
      chunks.push(bytes)
    }
    const size = chunks.reduce((total, chunk) => total + chunk.byteLength, 0)
    const joined = new Uint8Array(size)
    let offset = 0
    for (const chunk of chunks) {
      joined.set(chunk, offset)
      offset += chunk.byteLength
    }
    tracked.set(this, joined)
  }
}

function base64(bytes) {
  let binary = ''
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000))
  }
  return btoa(binary)
}

globalThis.Blob = TrackedBlob
URL.createObjectURL = (object) => {
  const bytes = tracked.get(object)
  if (!bytes) return nativeCreate(object)
  return `data:${object.type || 'application/octet-stream'};base64,${base64(bytes)}`
}
URL.revokeObjectURL = (url) => {
  if (typeof url === 'string' && !url.startsWith('data:')) nativeRevoke(url)
}

// Bitmaps the renderer draws itself (EMF previews, 3D shading) reach it through toBlob.
// Encode them through toDataURL so the resulting Blob is a tracked one.
HTMLCanvasElement.prototype.toBlob = function toBlob(callback, type = 'image/png', quality) {
  const url = this.toDataURL(type, quality)
  const comma = url.indexOf(',')
  const binary = atob(url.slice(comma + 1))
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  const mime = url.slice(5, url.indexOf(';'))
  setTimeout(() => callback(new TrackedBlob([bytes], { type: mime })), 0)
}
