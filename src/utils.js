/**
 * 将 ArrayBuffer 转换为 Base64 编码字符串
 * @param {ArrayBuffer} arrayBuffer - 要转换的 ArrayBuffer
 * @returns {string} Base64 编码的字符串
 */
export function base64ArrayBuffer(arrayBuffer) {
  // Base64 编码字符集
  const encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
  const bytes = new Uint8Array(arrayBuffer)
  const byteLength = bytes.byteLength
  const byteRemainder = byteLength % 3
  const mainLength = byteLength - byteRemainder
  
  let base64 = ''
  let a, b, c, d
  let chunk

  // 处理主要的字节（每3个字节一组）
  for (let i = 0; i < mainLength; i = i + 3) {
    // 将3个字节合并为一个24位的数
    chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2]
    // 将24位数分割为4个6位的数，并转换为Base64字符
    a = (chunk & 16515072) >> 18
    b = (chunk & 258048) >> 12
    c = (chunk & 4032) >> 6
    d = chunk & 63
    base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d]
  }

  // 处理剩余的字节（1或2个字节）
  if (byteRemainder === 1) {
    chunk = bytes[mainLength]
    a = (chunk & 252) >> 2
    b = (chunk & 3) << 4
    base64 += encodings[a] + encodings[b] + '=='
  } 
  else if (byteRemainder === 2) {
    chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1]
    a = (chunk & 64512) >> 10
    b = (chunk & 1008) >> 4
    c = (chunk & 15) << 2
    base64 += encodings[a] + encodings[b] + encodings[c] + '='
  }

  return base64
}

/**
 * 提取文件扩展名
 * @param {string} filename - 文件名
 * @returns {string} 文件扩展名
 */
export function extractFileExtension(filename) {
  return filename.substr((~-filename.lastIndexOf('.') >>> 0) + 2)
}

/**
 * 对节点或数组的每个元素执行函数
 * @param {Array|*} node - 要处理的节点或数组
 * @param {Function} func - 要执行的函数
 * @returns {string} 处理结果
 */
export function eachElement(node, func) {
  if (!node) return node

  let result = ''
  if (node.constructor === Array) {
    for (let i = 0; i < node.length; i++) {
      result += func(node[i], i)
    }
  } 
  else result += func(node, 0)

  return result
}

/**
 * 根据路径列表获取对象中的文本
 * @param {Object} node - 要处理的对象
 * @param {Array} path - 路径列表
 * @returns {*} 获取到的文本或节点
 * @throws {Error} 如果路径不是数组类型
 */
export function getTextByPathList(node, path) {
  if (path.constructor !== Array) throw Error('Error of path type! path is not array.')

  if (!node) return node

  for (const key of path) {
    node = node[key]
    if (!node) return node
  }

  return node
}

/**
 * 将角度转换为度数（60000单位等于1度）
 * @param {number} angle - 要转换的角度
 * @returns {number} 转换后的度数
 */
export function angleToDegrees(angle) {
  if (!angle) return 0
  return Math.round(angle / 60000)
}

/**
 * 转义HTML特殊字符
 * @param {string} text - 要转义的文本
 * @returns {string} 转义后的文本
 */
export function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;',
  }
  return text.replace(/[&<>"']/g, m => map[m])
}

/**
 * 根据文件扩展名获取MIME类型
 * @param {string} imgFileExt - 文件扩展名
 * @returns {string} MIME类型
 */
export function getMimeType(imgFileExt) {
  let mimeType = ''
  switch (imgFileExt.toLowerCase()) {
    case 'jpg':
    case 'jpeg':
      mimeType = 'image/jpeg'
      break
    case 'png':
      mimeType = 'image/png'
      break
    case 'gif':
      mimeType = 'image/gif'
      break
    case 'emf':
      mimeType = 'image/x-emf'
      break
    case 'wmf':
      mimeType = 'image/x-wmf'
      break
    case 'svg':
      mimeType = 'image/svg+xml'
      break
    case 'mp4':
      mimeType = 'video/mp4'
      break
    case 'webm':
      mimeType = 'video/webm'
      break
    case 'ogg':
      mimeType = 'video/ogg'
      break
    case 'avi':
      mimeType = 'video/avi'
      break
    case 'mpg':
      mimeType = 'video/mpg'
      break
    case 'wmv':
      mimeType = 'video/wmv'
      break
    case 'mp3':
      mimeType = 'audio/mpeg'
      break
    case 'wav':
      mimeType = 'audio/wav'
      break
    case 'tif':
      mimeType = 'image/tiff'
      break
    case 'tiff':
      mimeType = 'image/tiff'
      break
    default:
  }
  return mimeType
}

/**
 * 检查是否为有效的视频链接
 * @param {string} vdoFile - 要检查的URL
 * @returns {boolean} 是否为有效的视频链接
 */
export function isVideoLink(vdoFile) {
  const urlRegex = /^(https?|ftp):\/\/([a-zA-Z0-9.-]+(:[a-zA-Z0-9.&%$-]+)*@)*((25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])){3}|([a-zA-Z0-9-]+\.)*[a-zA-Z0-9-]+\.(com|edu|gov|int|mil|net|org|biz|arpa|info|name|pro|aero|coop|museum|[a-zA-Z]{2}))(:[0-9]+)*(\/($|[a-zA-Z0-9.,?'\\+&%$#=~_-]+))*$/
  return urlRegex.test(vdoFile)
}

/**
 * 将数字转换为两位十六进制字符串
 * @param {number} n - 要转换的数字
 * @returns {string} 两位十六进制字符串
 */
export function toHex(n) {
  let hex = n.toString(16)
  while (hex.length < 2) {
    hex = '0' + hex
  }
  return hex
}
