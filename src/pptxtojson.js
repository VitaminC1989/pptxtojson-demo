import JSZip from 'jszip'
import { readXmlFile } from './readXmlFile'
import { getBorder } from './border'
import { getSlideBackgroundFill, getShapeFill, getSolidFill } from './fill'
import { getChartInfo } from './chart'
import { getVerticalAlign } from './align'
import { getPosition, getSize } from './position'
import { genTextBody } from './text'
import { getCustomShapePath } from './shape'
import { extractFileExtension, base64ArrayBuffer, getTextByPathList, angleToDegrees, getMimeType, isVideoLink, escapeHtml } from './utils'
import { getShadow } from './shadow'
import { getTableBorders, getTableCellParams, getTableRowParams } from './table'
import { RATIO_EMUs_Points } from './constants'

/**
 * 解析PPTX文件，提取幻灯片信息和内容。
 * @param {File} file - 需要解析的PPTX文件。
 * @returns {Promise<{slides: any[], size: {width: number, height: number}}>} - 解析后的结果，包含所有幻灯片的信息和PPTX文件的大小。
 */
export async function parse(file) {
  const slides = [] // 用于存储所有解析后的幻灯片信息
  
  const zip = await JSZip.loadAsync(file) // 使用JSZip库异步加载PPTX文件

  const filesInfo = await getContentTypes(zip) // 获取PPTX文件中的内容类型信息
  const { width, height, defaultTextStyle } = await getSlideInfo(zip) // 获取幻灯片的大小和默认文本样式信息
  const themeContent = await loadTheme(zip) // 加载PPTX文件的主题信息

  console.log('zip', zip)
  console.log('filesInfo', filesInfo)
  console.log('width', width)
  console.log('height', height)
  console.log('defaultTextStyle', defaultTextStyle)
  console.log('themeContent', themeContent)

  for (const filename of filesInfo.slides) { // 遍历所有幻灯片文件
    const singleSlide = await processSingleSlide(zip, filename, themeContent, defaultTextStyle) // 解析单个幻灯片
    slides.push(singleSlide) // 将解析后的幻灯片信息添加到数组中
  }

  return {
    slides, // 解析后的所有幻灯片信息
    size: {
      width, // PPTX文件的宽度
      height, // PPTX文件的高度
    },
  }
}

/**
 * 获取PPTX文件内容类型
 * @param {JSZip} zip - JSZip实例
 * @returns {Promise<{slides: string[], slideLayouts: string[]}>} - 幻灯片和幻灯片布局的文件路径
 */
async function getContentTypes(zip) {
  // 读取[Content_Types].xml文件
  const ContentTypesJson = await readXmlFile(zip, '[Content_Types].xml')
  const subObj = ContentTypesJson['Types']['Override']
  let slidesLocArray = []
  let slideLayoutsLocArray = []

  // 遍历所有Override元素，提取幻灯片和幻灯片布局的文件路径
  for (const item of subObj) {
    switch (item['attrs']['ContentType']) {
      case 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml':
        slidesLocArray.push(item['attrs']['PartName'].substr(1))
        break
      case 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml':
        slideLayoutsLocArray.push(item['attrs']['PartName'].substr(1))
        break
      default:
    }
  }
  
  // 定义排序函数，按照幻灯片编号排序
  const sortSlideXml = (p1, p2) => {
    const n1 = +/(\d+)\.xml/.exec(p1)[1]
    const n2 = +/(\d+)\.xml/.exec(p2)[1]
    return n1 - n2
  }
  slidesLocArray = slidesLocArray.sort(sortSlideXml)
  slideLayoutsLocArray = slideLayoutsLocArray.sort(sortSlideXml)
  
  return {
    slides: slidesLocArray,
    slideLayouts: slideLayoutsLocArray,
  }
}

/**
 * 获取幻灯片信息
 * @param {JSZip} zip - JSZip实例
 * @returns {Promise<{width: number, height: number, defaultTextStyle: any}>} - 幻灯片大小和默认文本样式
 */
async function getSlideInfo(zip) {
  // 读取presentation.xml文件
  const content = await readXmlFile(zip, 'ppt/presentation.xml')
  const sldSzAttrs = content['p:presentation']['p:sldSz']['attrs']
  const defaultTextStyle = content['p:presentation']['p:defaultTextStyle']
  return {
    width: parseInt(sldSzAttrs['cx']) * RATIO_EMUs_Points,
    height: parseInt(sldSzAttrs['cy']) * RATIO_EMUs_Points,
    defaultTextStyle,
  }
}

/**
 * 加载主题
 * @param {JSZip} zip - JSZip实例
 * @returns {Promise<any>} - 主题内容
 */
async function loadTheme(zip) {
  // 读取presentation.xml.rels文件
  const preResContent = await readXmlFile(zip, 'ppt/_rels/presentation.xml.rels')
  const relationshipArray = preResContent['Relationships']['Relationship']
  let themeURI

  // 查找主题文件的URI
  if (relationshipArray.constructor === Array) {
    for (const relationshipItem of relationshipArray) {
      if (relationshipItem['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
        themeURI = relationshipItem['attrs']['Target']
        break
      }
    }
  } 
  else if (relationshipArray['attrs']['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme') {
    themeURI = relationshipArray['attrs']['Target']
  }
  if (!themeURI) throw Error(`Can't open theme file.`)

  // 读取主题文件内容
  return await readXmlFile(zip, 'ppt/' + themeURI)
}

/**
 * 处理单个幻灯片
 * @param {JSZip} zip - JSZip实例
 * @param {string} sldFileName - 幻灯片文件名
 * @param {any} themeContent - 主题内容
 * @param {any} defaultTextStyle - 默认文本样式
 * @returns {Promise<{fill: string, elements: any[]}>} - 处理后的幻灯片信息
 */
async function processSingleSlide(zip, sldFileName, themeContent, defaultTextStyle) {
  // 读取幻灯片关系文件
  const resName = sldFileName.replace('slides/slide', 'slides/_rels/slide') + '.rels'
  const resContent = await readXmlFile(zip, resName)
  let relationshipArray = resContent['Relationships']['Relationship']
  let layoutFilename = ''
  let diagramFilename = ''
  const slideResObj = {}

  // 处理幻灯片关系
  if (relationshipArray.constructor === Array) {
    for (const relationshipArrayItem of relationshipArray) {
      switch (relationshipArrayItem['attrs']['Type']) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout':
          layoutFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          break
        case 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing':
          diagramFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          slideResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          }
          break
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart':
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink':
        default:
          slideResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
      }
    }
  } 
  else layoutFilename = relationshipArray['attrs']['Target'].replace('../', 'ppt/')

  // 读取幻灯片布局内容
  const slideLayoutContent = await readXmlFile(zip, layoutFilename)
  const slideLayoutTables = await indexNodes(slideLayoutContent)

  // 读取幻灯片布局关系文件
  const slideLayoutResFilename = layoutFilename.replace('slideLayouts/slideLayout', 'slideLayouts/_rels/slideLayout') + '.rels'
  const slideLayoutResContent = await readXmlFile(zip, slideLayoutResFilename)
  relationshipArray = slideLayoutResContent['Relationships']['Relationship']

  let masterFilename = ''
  const layoutResObj = {}
  if (relationshipArray.constructor === Array) {
    for (const relationshipArrayItem of relationshipArray) {
      switch (relationshipArrayItem['attrs']['Type']) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster':
          masterFilename = relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          break
        default:
          layoutResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
      }
    }
  } 
  else masterFilename = relationshipArray['attrs']['Target'].replace('../', 'ppt/')

  // 读取幻灯片母版内容
  const slideMasterContent = await readXmlFile(zip, masterFilename)
  const slideMasterTextStyles = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:txStyles'])
  const slideMasterTables = indexNodes(slideMasterContent)

  // 读取幻灯片母版关系文件
  const slideMasterResFilename = masterFilename.replace('slideMasters/slideMaster', 'slideMasters/_rels/slideMaster') + '.rels'
  const slideMasterResContent = await readXmlFile(zip, slideMasterResFilename)
  relationshipArray = slideMasterResContent['Relationships']['Relationship']

  let themeFilename = ''
  const masterResObj = {}
  if (relationshipArray.constructor === Array) {
    for (const relationshipArrayItem of relationshipArray) {
      switch (relationshipArrayItem['attrs']['Type']) {
        case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme':
          break
        default:
          masterResObj[relationshipArrayItem['attrs']['Id']] = {
            type: relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            target: relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/'),
          }
      }
    }
  }
  else themeFilename = relationshipArray['attrs']['Target'].replace('../', 'ppt/')

  // 处理主题关系
  const themeResObj = {}
  if (themeFilename) {
    const themeName = themeFilename.split('/').pop()
    const themeResFileName = themeFilename.replace(themeName, '_rels/' + themeName) + '.rels'
    const themeResContent = await readXmlFile(zip, themeResFileName)
    if (themeResContent) {
      relationshipArray = themeResContent['Relationships']['Relationship']
      if (relationshipArray) {
        if (relationshipArray.constructor === Array) {
          for (const relationshipArrayItem of relationshipArray) {
            themeResObj[relationshipArrayItem['attrs']['Id']] = {
              'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
              'target': relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
            }
          }
        } 
        else {
          themeResObj[relationshipArray['attrs']['Id']] = {
            'type': relationshipArray['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            'target': relationshipArray['attrs']['Target'].replace('../', 'ppt/')
          }
        }
      }
    }
  }

  // 处理图表关系
  const diagramResObj = {}
  let digramFileContent = {}
  if (diagramFilename) {
    const diagName = diagramFilename.split('/').pop()
    const diagramResFileName = diagramFilename.replace(diagName, '_rels/' + diagName) + '.rels'
    digramFileContent = await readXmlFile(zip, diagramFilename)
    if (digramFileContent && digramFileContent && digramFileContent) {
      let digramFileContentObjToStr = JSON.stringify(digramFileContent)
      digramFileContentObjToStr = digramFileContentObjToStr.replace(/dsp:/g, 'p:')
      digramFileContent = JSON.parse(digramFileContentObjToStr)
    }
    const digramResContent = await readXmlFile(zip, diagramResFileName)
    if (digramResContent) {
      relationshipArray = digramResContent['Relationships']['Relationship']
      if (relationshipArray.constructor === Array) {
        for (const relationshipArrayItem of relationshipArray) {
          diagramResObj[relationshipArrayItem['attrs']['Id']] = {
            'type': relationshipArrayItem['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
            'target': relationshipArrayItem['attrs']['Target'].replace('../', 'ppt/')
          }
        }
      } 
      else {
        diagramResObj[relationshipArray['attrs']['Id']] = {
          'type': relationshipArray['attrs']['Type'].replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/', ''),
          'target': relationshipArray['attrs']['Target'].replace('../', 'ppt/')
        }
      }
    }
  }

  // 读取表格样式
  const tableStyles = await readXmlFile(zip, 'ppt/tableStyles.xml')

  // 读取幻灯片内容
  const slideContent = await readXmlFile(zip, sldFileName)
  const nodes = slideContent['p:sld']['p:cSld']['p:spTree']
  const warpObj = {
    zip,
    slideLayoutContent,
    slideLayoutTables,
    slideMasterContent,
    slideMasterTables,
    slideContent,
    tableStyles,
    slideResObj,
    slideMasterTextStyles,
    layoutResObj,
    masterResObj,
    themeContent,
    themeResObj,
    digramFileContent,
    diagramResObj,
    defaultTextStyle,
  }
  // const bgElements = await getBackground(warpObj)
  const bgColor = await getSlideBackgroundFill(warpObj)

  // 处理幻灯片中的所有元素
  const elements = []
  for (const nodeKey in nodes) {
    if (nodes[nodeKey].constructor === Array) {
      for (const node of nodes[nodeKey]) {
        const ret = await processNodesInSlide(nodeKey, node, warpObj, 'slide')
        if (ret) elements.push(ret)
      }
    } 
    else {
      const ret = await processNodesInSlide(nodeKey, nodes[nodeKey], warpObj, 'slide')
      if (ret) elements.push(ret)
    }
  }

  return {
    fill: bgColor,
    elements,
  }
}

// async function getBackground(warpObj) {
//   const elements = []
//   const slideLayoutContent = warpObj['slideLayoutContent']
//   const slideMasterContent = warpObj['slideMasterContent']
//   const nodesSldLayout = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:spTree'])
//   const nodesSldMaster = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:spTree'])

//   const showMasterSp = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'attrs', 'showMasterSp'])
//   if (nodesSldLayout) {
//     for (const nodeKey in nodesSldLayout) {
//       if (nodesSldLayout[nodeKey].constructor === Array) {
//         for (let i = 0; i < nodesSldLayout[nodeKey].length; i++) {
//           const ph_type = getTextByPathList(nodesSldLayout[nodeKey][i], ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
//           if (ph_type !== 'pic') {
//             const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey][i], warpObj, 'slideLayoutBg')
//             if (ret) elements.push(ret)
//           }
//         }
//       } 
//       else {
//         const ph_type = getTextByPathList(nodesSldLayout[nodeKey], ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
//         if (ph_type !== 'pic') {
//           const ret = await processNodesInSlide(nodeKey, nodesSldLayout[nodeKey], warpObj, 'slideLayoutBg')
//           if (ret) elements.push(ret)
//         }
//       }
//     }
//   }
//   if (nodesSldMaster && (showMasterSp === '1' || showMasterSp)) {
//     for (const nodeKey in nodesSldMaster) {
//       if (nodesSldMaster[nodeKey].constructor === Array) {
//         for (let i = 0; i < nodesSldMaster[nodeKey].length; i++) {
//           const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey][i], warpObj, 'slideMasterBg')
//           if (ret) elements.push(ret)
//         }
//       } 
//       else {
//         const ret = await processNodesInSlide(nodeKey, nodesSldMaster[nodeKey], warpObj, 'slideMasterBg')
//         if (ret) elements.push(ret)
//       }
//     }
//   }
//   return elements
// }

/**
 * 这个函数用于在内容中索引节点。
 * @param {Object} content - 要索引的内容。
 * @returns {Object} - 索引的节点。
 */
function indexNodes(content) {
  const keys = Object.keys(content)
  const spTreeNode = content[keys[0]]['p:cSld']['p:spTree']
  const idTable = {}
  const idxTable = {}
  const typeTable = {}

  for (const key in spTreeNode) {
    if (key === 'p:nvGrpSpPr' || key === 'p:grpSpPr') continue

    const targetNode = spTreeNode[key]

    if (targetNode.constructor === Array) {
      for (const targetNodeItem of targetNode) {
        const nvSpPrNode = targetNodeItem['p:nvSpPr']
        const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
        const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
        const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

        if (id) idTable[id] = targetNodeItem
        if (idx) idxTable[idx] = targetNodeItem
        if (type) typeTable[type] = targetNodeItem
      }
    } 
    else {
      const nvSpPrNode = targetNode['p:nvSpPr']
      const id = getTextByPathList(nvSpPrNode, ['p:cNvPr', 'attrs', 'id'])
      const idx = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'idx'])
      const type = getTextByPathList(nvSpPrNode, ['p:nvPr', 'p:ph', 'attrs', 'type'])

      if (id) idTable[id] = targetNode
      if (idx) idxTable[idx] = targetNode
      if (type) typeTable[type] = targetNode
    }
  }

  return { idTable, idxTable, typeTable }
}

/**
 * 这个函数用于处理幻灯片中的节点。
 * @param {string} nodeKey - 节点的键。
 * @param {Object} nodeValue - 节点的值。
 * @param {Object} warpObj - 包含幻灯片内容的对象。
 * @param {string} source - 节点的来源。
 * @returns {Object} - 处理后的节点。
 */
async function processNodesInSlide(nodeKey, nodeValue, warpObj, source) {
  let json

  switch (nodeKey) {
    case 'p:sp': // Shape, Text
      json = processSpNode(nodeValue, warpObj, source)
      console.log('[processNodesInSlide] p:sp', json)
      break
    case 'p:cxnSp': // Shape, Text
      json = processCxnSpNode(nodeValue, warpObj, source)
      console.log('[processNodesInSlide] p:cxnSp')
      break
    case 'p:pic': // Image, Video, Audio
      json = processPicNode(nodeValue, warpObj, source)
      break
    case 'p:graphicFrame': // Chart, Diagram, Table
      json = await processGraphicFrameNode(nodeValue, warpObj, source)
      break
    case 'p:grpSp':
      json = await processGroupSpNode(nodeValue, warpObj, source)
      break
    case 'mc:AlternateContent':
      json = await processGroupSpNode(getTextByPathList(nodeValue, ['mc:Fallback']), warpObj, source)
      break
    default:
  }

  return json
}

/**
 * 这个函数用于处理组合节点。
 * @param {Object} node - 组合节点。
 * @param {Object} warpObj - 包含幻灯片内容的对象。
 * @param {string} source - 节点的来源。
 * @returns {Object} - 处理后的组合节点。
 */
async function processGroupSpNode(node, warpObj, source) {
  const xfrmNode = getTextByPathList(node, ['p:grpSpPr', 'a:xfrm'])
  if (!xfrmNode) return null

  const x = parseInt(xfrmNode['a:off']['attrs']['x']) * RATIO_EMUs_Points
  const y = parseInt(xfrmNode['a:off']['attrs']['y']) * RATIO_EMUs_Points
  const chx = parseInt(xfrmNode['a:chOff']['attrs']['x']) * RATIO_EMUs_Points
  const chy = parseInt(xfrmNode['a:chOff']['attrs']['y']) * RATIO_EMUs_Points
  const cx = parseInt(xfrmNode['a:ext']['attrs']['cx']) * RATIO_EMUs_Points
  const cy = parseInt(xfrmNode['a:ext']['attrs']['cy']) * RATIO_EMUs_Points
  const chcx = parseInt(xfrmNode['a:chExt']['attrs']['cx']) * RATIO_EMUs_Points
  const chcy = parseInt(xfrmNode['a:chExt']['attrs']['cy']) * RATIO_EMUs_Points

  let rotate = getTextByPathList(xfrmNode, ['attrs', 'rot']) || 0
  if (rotate) rotate = angleToDegrees(rotate)

  const ws = cx / chcx
  const hs = cy / chcy

  const elements = []
  for (const nodeKey in node) {
    if (node[nodeKey].constructor === Array) {
      for (const item of node[nodeKey]) {
        const ret = await processNodesInSlide(nodeKey, item, warpObj, source)
        if (ret) elements.push(ret)
      }
    }
    else {
      const ret = await processNodesInSlide(nodeKey, node[nodeKey], warpObj, source)
      if (ret) elements.push(ret)
    }
  }

  return {
    type: 'group',
    top: y,
    left: x,
    width: cx,
    height: cy,
    rotate,
    elements: elements.map(element => ({
      ...element,
      left: (element.left - chx) * ws,
      top: (element.top - chy) * hs,
      width: element.width * ws,
      height: element.height * hs,
    }))
  }
}

/**
 * 处理幻灯片中的形状节点
 * @param {Object} node - 形状节点对象
 * @param {Object} warpObj - 包含幻灯片内容的包装对象
 * @param {string} source - 节点来源
 * @returns {Object} 处理后的形状对象
 */
function processSpNode(node, warpObj, source) {
  console.log('[processSpNode] node', node, warpObj.fillColor)
  // 获取形状名称
  const name = getTextByPathList(node, ['p:nvSpPr', 'p:cNvPr', 'attrs', 'name'])
  
  // 获取占位符索引
  const idx = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'idx'])
  
  // 获取占位符类型
  let type = getTextByPathList(node, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  let slideLayoutSpNode, slideMasterSpNode

  // 根据类型和索引查找对应的布局和母版节点
  if (type) {
    if (idx) {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    } 
    else {
      slideLayoutSpNode = warpObj['slideLayoutTables']['typeTable'][type]
      slideMasterSpNode = warpObj['slideMasterTables']['typeTable'][type]
    }
  }
  else if (idx) {
    slideLayoutSpNode = warpObj['slideLayoutTables']['idxTable'][idx]
    slideMasterSpNode = warpObj['slideMasterTables']['idxTable'][idx]
  }

  // 如果没有类型,尝试确定类型
  if (!type) {
    // 检查是否为文本框
    const txBoxVal = getTextByPathList(node, ['p:nvSpPr', 'p:cNvSpPr', 'attrs', 'txBox'])
    if (txBoxVal === '1') type = 'text'
  }
  // 从布局或母版中获取类型
  if (!type) type = getTextByPathList(slideLayoutSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])
  if (!type) type = getTextByPathList(slideMasterSpNode, ['p:nvSpPr', 'p:nvPr', 'p:ph', 'attrs', 'type'])

  // 如果仍然没有类型,根据来源设置默认类型
  if (!type) {
    if (source === 'diagramBg') type = 'diagram'
    else type = 'obj'
  }

  // 生成并返回形状对象
  return genShape(node, slideLayoutSpNode, slideMasterSpNode, name, type, warpObj)
}

function processCxnSpNode(node, warpObj) {
  const name = node['p:nvCxnSpPr']['p:cNvPr']['attrs']['name']
  const type = (node['p:nvCxnSpPr']['p:nvPr']['p:ph'] === undefined) ? undefined : node['p:nvSpPr']['p:nvPr']['p:ph']['attrs']['type']

  return genShape(node, undefined, undefined, name, type, warpObj)
}

/**
 * 生成形状对象
 * @param {Object} node - 当前节点
 * @param {Object} slideLayoutSpNode - 幻灯片布局节点
 * @param {Object} slideMasterSpNode - 幻灯片母版节点
 * @param {string} name - 形状名称
 * @param {string} type - 形状类型
 * @param {Object} warpObj - 包含幻灯片内容的包装对象
 * @returns {Object} 生成的形状对象
 */
async function genShape(node, slideLayoutSpNode, slideMasterSpNode, name, type, warpObj) {
  console.log('[genShape] ---name---', name)
  // 获取变换信息
  const xfrmList = ['p:spPr', 'a:xfrm']
  const slideXfrmNode = getTextByPathList(node, xfrmList)
  const slideLayoutXfrmNode = getTextByPathList(slideLayoutSpNode, xfrmList)
  const slideMasterXfrmNode = getTextByPathList(slideMasterSpNode, xfrmList)

  // 获取形状类型
  const shapType = getTextByPathList(node, ['p:spPr', 'a:prstGeom', 'attrs', 'prst'])
  console.log('[genShape] shapType', shapType)
  const custShapType = getTextByPathList(node, ['p:spPr', 'a:custGeom'])

  // 获取位置和大小
  const { top, left } = getPosition(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)
  const { width, height } = getSize(slideXfrmNode, slideLayoutXfrmNode, slideMasterXfrmNode)

  // 获取翻转信息
  const isFlipV = getTextByPathList(slideXfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(slideXfrmNode, ['attrs', 'flipH']) === '1'

  // 获取旋转角度
  const rotate = angleToDegrees(getTextByPathList(slideXfrmNode, ['attrs', 'rot']))

  // 获取文本旋转角度
  const txtXframeNode = getTextByPathList(node, ['p:txXfrm'])
  let txtRotate
  if (txtXframeNode) {
    const txtXframeRot = getTextByPathList(txtXframeNode, ['attrs', 'rot'])
    if (txtXframeRot) txtRotate = angleToDegrees(txtXframeRot) + 90
  } 
  else txtRotate = rotate

  // 生成文本内容
  let content = ''
  if (node['p:txBody']) content = genTextBody(node['p:txBody'], node, slideLayoutSpNode, type, warpObj)

  // 获取边框信息
  const { borderColor, borderWidth, borderType, strokeDasharray } = getBorder(node, type, warpObj)
  // 获取填充颜色
  const fillColor = getShapeFill(node, undefined, warpObj) || ''
  // 获取渐变色
  const gradientFill = await getSlideBackgroundFill(warpObj) || ''
  console.log('[genShape] gradientFill', gradientFill)

  // 获取阴影信息
  let shadow
  const outerShdwNode = getTextByPathList(node, ['p:spPr', 'a:effectLst', 'a:outerShdw'])
  if (outerShdwNode) shadow = getShadow(outerShdwNode, warpObj)

  // 获取垂直对齐方式
  const vAlign = getVerticalAlign(node, slideLayoutSpNode, slideMasterSpNode, type)
  // 判断是否为垂直文本
  const isVertical = getTextByPathList(node, ['p:txBody', 'a:bodyPr', 'attrs', 'vert']) === 'eaVert'

  // 构建基本数据对象
  const data = {
    left,
    top,
    width,
    height,
    borderColor,
    borderWidth,
    borderType,
    borderStrokeDasharray: strokeDasharray,
    fillColor,
    content,
    isFlipV,
    isFlipH,
    rotate,
    vAlign,
    name,
  }

  // 添加阴影信息（如果存在）
  if (shadow) data.shadow = shadow

  // 处理自定义形状
  if (custShapType && type !== 'diagram') {
    const ext = getTextByPathList(slideXfrmNode, ['a:ext', 'attrs'])
    const w = parseInt(ext['cx']) * RATIO_EMUs_Points
    const h = parseInt(ext['cy']) * RATIO_EMUs_Points
    const d = getCustomShapePath(custShapType, w, h)

    return {
      ...data,
      type: 'shape',
      shapType: 'custom',
      path: d,
    }
  }
  // 处理预设形状
  if (shapType && (type === 'obj' || !type)) {
    return {
      ...data,
      type: 'shape',
      shapType,
    }
  }
  // 处理文本形状
  return {
    ...data,
    type: 'text',
    isVertical,
    rotate: txtRotate,
  }
}

async function processPicNode(node, warpObj, source) {
  let resObj
  if (source === 'slideMasterBg') resObj = warpObj['masterResObj']
  else if (source === 'slideLayoutBg') resObj = warpObj['layoutResObj']
  else resObj = warpObj['slideResObj']
  
  const rid = node['p:blipFill']['a:blip']['attrs']['r:embed']
  const imgName = resObj[rid]['target']
  const imgFileExt = extractFileExtension(imgName).toLowerCase()
  const zip = warpObj['zip']
  const imgArrayBuffer = await zip.file(imgName).async('arraybuffer')
  const xfrmNode = node['p:spPr']['a:xfrm']

  const mimeType = getMimeType(imgFileExt)
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)
  const src = `data:${mimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`

  const isFlipV = getTextByPathList(xfrmNode, ['attrs', 'flipV']) === '1'
  const isFlipH = getTextByPathList(xfrmNode, ['attrs', 'flipH']) === '1'

  let rotate = 0
  const rotateNode = getTextByPathList(node, ['p:spPr', 'a:xfrm', 'attrs', 'rot'])
  if (rotateNode) rotate = angleToDegrees(rotateNode)

  const videoNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:videoFile'])
  let videoRid, videoFile, videoFileExt, videoMimeType, uInt8ArrayVideo, videoBlob
  let isVdeoLink = false

  if (videoNode) {
    videoRid = videoNode['attrs']['r:link']
    videoFile = resObj[videoRid]['target']
    if (isVideoLink(videoFile)) {
      videoFile = escapeHtml(videoFile)
      isVdeoLink = true
    } 
    else {
      videoFileExt = extractFileExtension(videoFile).toLowerCase()
      if (videoFileExt === 'mp4' || videoFileExt === 'webm' || videoFileExt === 'ogg') {
        uInt8ArrayVideo = await zip.file(videoFile).async('arraybuffer')
        videoMimeType = getMimeType(videoFileExt)
        videoBlob = URL.createObjectURL(new Blob([uInt8ArrayVideo], {
          type: videoMimeType
        }))
      }
    }
  }

  const audioNode = getTextByPathList(node, ['p:nvPicPr', 'p:nvPr', 'a:audioFile'])
  let audioRid, audioFile, audioFileExt, uInt8ArrayAudio, audioBlob
  if (audioNode) {
    audioRid = audioNode['attrs']['r:link']
    audioFile = resObj[audioRid]['target']
    audioFileExt = extractFileExtension(audioFile).toLowerCase()
    if (audioFileExt === 'mp3' || audioFileExt === 'wav' || audioFileExt === 'ogg') {
      uInt8ArrayAudio = await zip.file(audioFile).async('arraybuffer')
      audioBlob = URL.createObjectURL(new Blob([uInt8ArrayAudio]))
    }
  }

  if (videoNode && !isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width, 
      height,
      rotate,
      blob: videoBlob,
    }
  } 
  if (videoNode && isVdeoLink) {
    return {
      type: 'video',
      top,
      left,
      width, 
      height,
      rotate,
      src: videoFile,
    }
  }
  if (audioNode) {
    return {
      type: 'audio',
      top,
      left,
      width, 
      height,
      rotate,
      blob: audioBlob,
    }
  }
  return {
    type: 'image',
    top,
    left,
    width, 
    height,
    rotate,
    src,
    isFlipV,
    isFlipH
  }
}

async function processGraphicFrameNode(node, warpObj, source) {
  const graphicTypeUri = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'attrs', 'uri'])
  
  let result
  switch (graphicTypeUri) {
    case 'http://schemas.openxmlformats.org/drawingml/2006/table':
      result = genTable(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/chart':
      result = await genChart(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/drawingml/2006/diagram':
      result = genDiagram(node, warpObj)
      break
    case 'http://schemas.openxmlformats.org/presentationml/2006/ole':
      let oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'mc:AlternateContent', 'mc:Fallback', 'p:oleObj'])
      if (!oleObjNode) oleObjNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'p:oleObj'])
      else processGroupSpNode(oleObjNode, warpObj, source)
      break
    default:
  }
  return result
}

function genTable(node, warpObj) {
  const tableNode = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl'])
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const getTblPr = getTextByPathList(node, ['a:graphic', 'a:graphicData', 'a:tbl', 'a:tblPr'])

  const firstRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstRow'] : undefined
  const firstColAttr = getTblPr['attrs'] ? getTblPr['attrs']['firstCol'] : undefined
  const lastRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastRow'] : undefined
  const lastColAttr = getTblPr['attrs'] ? getTblPr['attrs']['lastCol'] : undefined
  const bandRowAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandRow'] : undefined
  const bandColAttr = getTblPr['attrs'] ? getTblPr['attrs']['bandCol'] : undefined
  const tblStylAttrObj = {
    isFrstRowAttr: (firstRowAttr && firstRowAttr === '1') ? 1 : 0,
    isFrstColAttr: (firstColAttr && firstColAttr === '1') ? 1 : 0,
    isLstRowAttr: (lastRowAttr && lastRowAttr === '1') ? 1 : 0,
    isLstColAttr: (lastColAttr && lastColAttr === '1') ? 1 : 0,
    isBandRowAttr: (bandRowAttr && bandRowAttr === '1') ? 1 : 0,
    isBandColAttr: (bandColAttr && bandColAttr === '1') ? 1 : 0,
  }

  let thisTblStyle
  const tbleStyleId = getTblPr['a:tableStyleId']
  if (tbleStyleId) {
    const tbleStylList = warpObj['tableStyles']['a:tblStyleLst']['a:tblStyle']
    if (tbleStylList) {
      if (tbleStylList.constructor === Array) {
        for (let k = 0; k < tbleStylList.length; k++) {
          if (tbleStylList[k]['attrs']['styleId'] === tbleStyleId) {
            thisTblStyle = tbleStylList[k]
          }
        }
      } 
      else {
        if (tbleStylList['attrs']['styleId'] === tbleStyleId) {
          thisTblStyle = tbleStylList
        }
      }
    }
  }
  if (thisTblStyle) thisTblStyle['tblStylAttrObj'] = tblStylAttrObj

  let tbl_border
  const tblStyl = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle'])
  const tblBorderStyl = getTextByPathList(tblStyl, ['a:tcBdr'])
  if (tblBorderStyl) {
    const tbl_borders = getTableBorders(tblBorderStyl, warpObj)
    if (tbl_borders) tbl_border = tbl_borders.bottom || tbl_borders.left || tbl_borders.right || tbl_borders.top
  }

  let tbl_bgcolor = ''
  let tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:tblBg', 'a:fillRef'])
  if (tbl_bgFillschemeClr) {
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }
  if (tbl_bgFillschemeClr === undefined) {
    tbl_bgFillschemeClr = getTextByPathList(thisTblStyle, ['a:wholeTbl', 'a:tcStyle', 'a:fill', 'a:solidFill'])
    tbl_bgcolor = getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj)
  }

  let trNodes = tableNode['a:tr']
  if (trNodes.constructor !== Array) trNodes = [trNodes]
  
  const data = []
  for (let i = 0; i < trNodes.length; i++) {
    const trNode = trNodes[i]

    const {
      fillColor,
      fontColor,
      fontBold,
    } = getTableRowParams(trNodes, i, tblStylAttrObj, thisTblStyle, warpObj)

    const tcNodes = trNode['a:tc']
    const tr = []

    if (tcNodes.constructor === Array) {
      for (let j = 0; j < tcNodes.length; j++) {
        const tcNode = tcNodes[j]
        let a_sorce
        if (j === 0 && tblStylAttrObj['isFrstColAttr'] === 1) {
          a_sorce = 'a:firstCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:seCell'])) {
            a_sorce = 'a:seCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 &&
            getTextByPathList(thisTblStyle, ['a:neCell'])) {
            a_sorce = 'a:neCell'
          }
        } 
        else if (
          (j > 0 && tblStylAttrObj['isBandColAttr'] === 1) &&
          !(tblStylAttrObj['isFrstColAttr'] === 1 && i === 0) &&
          !(tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1)) &&
          j !== (tcNodes.length - 1)
        ) {
          if ((j % 2) !== 0) {
            let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
            if (aBandNode === undefined) {
              aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
              if (aBandNode) a_sorce = 'a:band2V'
            } 
            else a_sorce = 'a:band2V'
          }
        }
        if (j === (tcNodes.length - 1) && tblStylAttrObj['isLstColAttr'] === 1) {
          a_sorce = 'a:lastCol'
          if (tblStylAttrObj['isLstRowAttr'] === 1 && i === (trNodes.length - 1) && getTextByPathList(thisTblStyle, ['a:swCell'])) {
            a_sorce = 'a:swCell'
          } 
          else if (tblStylAttrObj['isFrstRowAttr'] === 1 && i === 0 && getTextByPathList(thisTblStyle, ['a:nwCell'])) {
            a_sorce = 'a:nwCell'
          }
        }
        const text = genTextBody(tcNode['a:txBody'], tcNode, undefined, undefined, warpObj)
        const cell = getTableCellParams(tcNode, thisTblStyle, a_sorce, warpObj)
        const td = { text }
        if (cell.rowSpan) td.rowSpan = cell.rowSpan
        if (cell.colSpan) td.colSpan = cell.colSpan
        if (cell.vMerge) td.vMerge = cell.vMerge
        if (cell.hMerge) td.hMerge = cell.hMerge
        if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
        if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
        if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor

        tr.push(td)
      }
    } 
    else {
      let a_sorce
      if (tblStylAttrObj['isFrstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:firstCol'
      } 
      else if (tblStylAttrObj['isBandColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        let aBandNode = getTextByPathList(thisTblStyle, ['a:band2V'])
        if (!aBandNode) {
          aBandNode = getTextByPathList(thisTblStyle, ['a:band1V'])
          if (aBandNode) a_sorce = 'a:band2V'
        } 
        else a_sorce = 'a:band2V'
      }
      if (tblStylAttrObj['isLstColAttr'] === 1 && tblStylAttrObj['isLstRowAttr'] !== 1) {
        a_sorce = 'a:lastCol'
      }

      const text = genTextBody(tcNodes['a:txBody'], tcNodes, undefined, undefined, warpObj)
      const cell = getTableCellParams(tcNodes, thisTblStyle, a_sorce, warpObj)
      const td = { text }
      if (cell.rowSpan) td.rowSpan = cell.rowSpan
      if (cell.colSpan) td.colSpan = cell.colSpan
      if (cell.vMerge) td.vMerge = cell.vMerge
      if (cell.hMerge) td.hMerge = cell.hMerge
      if (cell.fontBold || fontBold) td.fontBold = cell.fontBold || fontBold
      if (cell.fontColor || fontColor) td.fontColor = cell.fontColor || fontColor
      if (cell.fillColor || fillColor || tbl_bgcolor) td.fillColor = cell.fillColor || fillColor || tbl_bgcolor

      tr.push(td)
    }
    data.push(tr)
  }

  return {
    type: 'table',
    top,
    left,
    width,
    height,
    data,
    ...(tbl_border || {}),
  }
}

async function genChart(node, warpObj) {
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { top, left } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)

  const rid = node['a:graphic']['a:graphicData']['c:chart']['attrs']['r:id']
  const refName = warpObj['slideResObj'][rid]['target']
  const content = await readXmlFile(warpObj['zip'], refName)
  const plotArea = getTextByPathList(content, ['c:chartSpace', 'c:chart', 'c:plotArea'])

  const chart = getChartInfo(plotArea)

  if (!chart) return {}

  const data = {
    type: 'chart',
    top,
    left,
    width,
    height,
    data: chart.data,
    chartType: chart.type,
  }
  if (chart.marker !== undefined) data.marker = chart.marker
  if (chart.barDir !== undefined) data.barDir = chart.barDir
  if (chart.holeSize !== undefined) data.holeSize = chart.holeSize
  if (chart.grouping !== undefined) data.grouping = chart.grouping
  if (chart.style !== undefined) data.style = chart.style

  return data
}

function genDiagram(node, warpObj) {
  const xfrmNode = getTextByPathList(node, ['p:xfrm'])
  const { left, top } = getPosition(xfrmNode, undefined, undefined)
  const { width, height } = getSize(xfrmNode, undefined, undefined)
  
  const dgmDrwSpArray = getTextByPathList(warpObj['digramFileContent'], ['p:drawing', 'p:spTree', 'p:sp'])
  const elements = []
  if (dgmDrwSpArray) {
    for (const item of dgmDrwSpArray) {
      const el = processSpNode(item, warpObj, 'diagramBg')
      if (el) elements.push(el)
    }
  }

  return {
    type: 'diagram',
    left,
    top,
    width,
    height,
    elements,
  }
}