export const MAX_DOCUMENT_BYTES = 64 * 1024 * 1024

/** Accepts only versioned binary documents delivered by Orca's frame API. */
export function parseDocument(doc) {
  if (!doc || !Number.isSafeInteger(doc.version)) return null
  if (doc.content?.kind !== 'bytes' || !(doc.content.bytes instanceof ArrayBuffer)) return null
  return { version: doc.version, bytes: doc.content.bytes }
}
