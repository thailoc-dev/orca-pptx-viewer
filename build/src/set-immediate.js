// JSZip imports the setImmediate polyfill for its side effect. The polyfill's string-callback
// fallback uses new Function, which the frame CSP refuses; JSZip passes only functions.
if (!globalThis.setImmediate) {
  const channel = new MessageChannel()
  const pending = new Map()
  let next = 0
  channel.port1.onmessage = ({ data }) => {
    const callback = pending.get(data)
    pending.delete(data)
    callback?.()
  }
  globalThis.setImmediate = (callback, ...args) => {
    if (typeof callback !== 'function') throw new TypeError('setImmediate expects a function')
    const id = ++next
    pending.set(id, () => callback(...args))
    channel.port2.postMessage(id)
    return id
  }
  globalThis.clearImmediate = (id) => pending.delete(id)
}
