import tinycolor from 'tinycolor2'
import { getSchemeColorFromTheme } from './schemeColor'
import {
  applyShade,
  applyTint,
  applyLumOff,
  applyLumMod,
  applyHueMod,
  applySatMod,
  hslToRgb,
  getColorName2Hex,
} from './color'

import {
  base64ArrayBuffer,
  getTextByPathList,
  angleToDegrees,
  escapeHtml,
  getMimeType,
  toHex,
} from './utils'

export function getFillType(node) {
  let fillType = ''
  if (node['a:noFill']) fillType = 'NO_FILL'
  if (node['a:solidFill']) fillType = 'SOLID_FILL'
  if (node['a:gradFill']) fillType = 'GRADIENT_FILL'
  if (node['a:pattFill']) fillType = 'PATTERN_FILL'
  if (node['a:blipFill']) fillType = 'PIC_FILL'
  if (node['a:grpFill']) fillType = 'GROUP_FILL'

  return fillType
}

export async function getPicFill(type, node, warpObj) {
  let img
  const rId = node['a:blip']['attrs']['r:embed']
  let imgPath
  if (type === 'slideBg' || type === 'slide') {
    imgPath = getTextByPathList(warpObj, ['slideResObj', rId, 'target'])
  }
  else if (type === 'slideLayoutBg') {
    imgPath = getTextByPathList(warpObj, ['layoutResObj', rId, 'target'])
  }
  else if (type === 'slideMasterBg') {
    imgPath = getTextByPathList(warpObj, ['masterResObj', rId, 'target'])
  }
  else if (type === 'themeBg') {
    imgPath = getTextByPathList(warpObj, ['themeResObj', rId, 'target'])
  }
  else if (type === 'diagramBg') {
    imgPath = getTextByPathList(warpObj, ['diagramResObj', rId, 'target'])
  }
  if (!imgPath) return imgPath

  img = getTextByPathList(warpObj, ['loaded-images', imgPath])
  if (!img) {
    imgPath = escapeHtml(imgPath)

    const imgExt = imgPath.split('.').pop()
    if (imgExt === 'xml') return undefined

    const imgArrayBuffer = await warpObj['zip'].file(imgPath).async('arraybuffer')
    const imgMimeType = getMimeType(imgExt)
    img = `data:${imgMimeType};base64,${base64ArrayBuffer(imgArrayBuffer)}`
  }
  return img
}

export async function getBgPicFill(bgPr, sorce, warpObj) {
  const picBase64 = await getPicFill(sorce, bgPr['a:blipFill'], warpObj)
  const aBlipNode = bgPr['a:blipFill']['a:blip']

  const aphaModFixNode = getTextByPathList(aBlipNode, ['a:alphaModFix', 'attrs'])
  let opacity = 1
  if (aphaModFixNode && aphaModFixNode['amt'] && aphaModFixNode['amt'] !== '') {
    opacity = parseInt(aphaModFixNode['amt']) / 100000
  }

  return {
    picBase64,
    opacity,
  }
}

export function getBgGradientFill(bgPr, phClr, slideMasterContent, warpObj) {
  if (bgPr) {
    const grdFill = bgPr['a:gradFill']
    const gsLst = grdFill['a:gsLst']['a:gs']
    const color_ary = []
    
    for (let i = 0; i < gsLst.length; i++) {
      const lo_color = getSolidFill(gsLst[i], slideMasterContent['p:sldMaster']['p:clrMap']['attrs'], phClr, warpObj)
      const pos = getTextByPathList(gsLst[i], ['attrs', 'pos'])

      color_ary[i] = {
        pos: pos ? (pos / 1000 + '%') : '',
        color: lo_color,
      }
    }
    const lin = grdFill['a:lin']
    let rot = 90
    if (lin) {
      rot = angleToDegrees(lin['attrs']['ang'])
      rot = rot + 90
    }

    return {
      rot,
      colors: color_ary.sort((a, b) => parseInt(a.pos) - parseInt(b.pos)),
    }
  }
  else if (phClr) {
    return phClr.indexOf('#') === -1 ? `#${phClr}` : phClr
  }
  return null
}

/**
 * 获取幻灯片背景填充
 * @param {Object} warpObj - 包含幻灯片内容、布局和母版的对象
 * @returns {Promise<Object>} 包含背景类型和值的对象
 */
export async function getSlideBackgroundFill(warpObj) {
  const slideContent = warpObj['slideContent']
  const slideLayoutContent = warpObj['slideLayoutContent']
  const slideMasterContent = warpObj['slideMasterContent']
  
  // 尝试从幻灯片获取背景属性或引用
  let bgPr = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:bg', 'p:bgPr'])
  let bgRef = getTextByPathList(slideContent, ['p:sld', 'p:cSld', 'p:bg', 'p:bgRef'])

  // 默认背景为白色
  let background = '#fff'
  let backgroundType = 'color'

  console.log('[getSlideBackgroundFill] bgPr', bgPr, bgRef, warpObj)

  if (bgPr) {
    // 如果存在背景属性,处理不同的填充类型
    const bgFillTyp = getFillType(bgPr)

    if (bgFillTyp === 'SOLID_FILL') {
      // 处理纯色填充
      const sldFill = bgPr['a:solidFill']
      let clrMapOvr
      // 获取颜色映射覆盖
      const sldClrMapOvr = getTextByPathList(slideContent, ['p:sld', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
      if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
      else {
        // 如果幻灯片没有颜色映射覆盖,尝试从布局或母版获取
        const sldClrMapOvr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:clrMapOvr', 'a:overrideClrMapping', 'attrs'])
        if (sldClrMapOvr) clrMapOvr = sldClrMapOvr
        else clrMapOvr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:clrMap', 'attrs'])
      }
      const sldBgClr = getSolidFill(sldFill, clrMapOvr, undefined, warpObj)
      background = sldBgClr
    }
    else if (bgFillTyp === 'GRADIENT_FILL') {
      // 处理渐变填充
      const gradientFill = getBgGradientFill(bgPr, undefined, slideMasterContent, warpObj)
      if (typeof gradientFill === 'string') {
        background = gradientFill
      }
      else if (gradientFill) {
        background = gradientFill
        backgroundType = 'gradient'
      }
    }
    else if (bgFillTyp === 'PIC_FILL') {
      // 处理图片填充
      background = await getBgPicFill(bgPr, 'slideBg', warpObj)
      backgroundType = 'image'
    }
  }
  else if (bgRef) {
    // 如果存在背景引用,处理引用的背景
    // ... (省略部分代码,处理逻辑类似)
  }
  else {
    // 如果幻灯片没有背景设置,尝试从布局获取
    bgPr = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:bg', 'p:bgPr'])
    bgRef = getTextByPathList(slideLayoutContent, ['p:sldLayout', 'p:cSld', 'p:bg', 'p:bgRef'])

    // ... (省略部分代码,处理逻辑类似)

    if (!bgPr && !bgRef) {
      // 如果布局也没有背景设置,尝试从母版获取
      bgPr = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:bg', 'p:bgPr'])
      bgRef = getTextByPathList(slideMasterContent, ['p:sldMaster', 'p:cSld', 'p:bg', 'p:bgRef'])

      // ... (省略部分代码,处理逻辑类似)
    }
  }

  // 返回背景类型和值
  return {
    type: backgroundType,
    value: background,
  }
}

/**
 * 获取形状的填充颜色
 * @param {Object} node - 包含形状信息的节点对象
 * @param {boolean} isSvgMode - 是否为SVG模式
 * @param {Object} warpObj - 包含主题和其他相关信息的对象
 * @returns {string} 填充颜色的十六进制值或'none'
 */
export function getShapeFill(node, isSvgMode, warpObj) {
  // 检查是否有noFill属性
  if (getTextByPathList(node, ['p:spPr', 'a:noFill'])) {
    return isSvgMode ? 'none' : ''
  }

  let fillColor

  // 尝试获取直接指定的RGB颜色
  if (!fillColor) {
    fillColor = getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:srgbClr', 'attrs', 'val'])
    console.log('1. 直接RGB颜色:', fillColor)
  }

  // 尝试从主题中获取方案颜色
  if (!fillColor) {
    const schemeClr = 'a:' + getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'attrs', 'val'])
    fillColor = getSchemeColorFromTheme(schemeClr, warpObj)
    console.log('2. 主题方案颜色:', fillColor)
  }

  // 尝试从填充引用中获取方案颜色
  if (!fillColor) {
    const schemeClr = 'a:' + getTextByPathList(node, ['p:style', 'a:fillRef', 'a:schemeClr', 'attrs', 'val'])
    fillColor = getSchemeColorFromTheme(schemeClr, warpObj)
    console.log('3. 填充引用颜色:', fillColor)
  }

  // 如果找到颜色,进行进一步处理
  if (fillColor) {
    fillColor = `#${fillColor}`
    console.log('4. 处理前的颜色:', fillColor)

    // 获取亮度调整参数
    let lumMod = parseInt(getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'a:lumMod', 'attrs', 'val'])) / 100000
    let lumOff = parseInt(getTextByPathList(node, ['p:spPr', 'a:solidFill', 'a:schemeClr', 'a:lumOff', 'attrs', 'val'])) / 100000
    if (isNaN(lumMod)) lumMod = 1.0
    if (isNaN(lumOff)) lumOff = 0

    // 应用亮度调整
    const color = tinycolor(fillColor).toHsl()
    const lum = color.l * lumMod + lumOff
    return tinycolor({ h: color.h, s: color.s, l: lum, a: color.a }).toHexString()
  } 

  // 如果是SVG模式且没有找到颜色,返回'none'
  if (isSvgMode) return 'none'

  // 返回找到的颜色或undefined
  return fillColor
}

export function getSolidFill(solidFill, clrMap, phClr, warpObj) {
  if (!solidFill) return ''

  let color = ''
  let clrNode

  if (solidFill['a:srgbClr']) {
    clrNode = solidFill['a:srgbClr']
    color = getTextByPathList(clrNode, ['attrs', 'val'])
  } 
  else if (solidFill['a:schemeClr']) {
    clrNode = solidFill['a:schemeClr']
    const schemeClr = 'a:' + getTextByPathList(clrNode, ['attrs', 'val'])
    color = getSchemeColorFromTheme(schemeClr, warpObj, clrMap, phClr) || ''
  }
  else if (solidFill['a:scrgbClr']) {
    clrNode = solidFill['a:scrgbClr']
    const defBultColorVals = clrNode['attrs']
    const red = (defBultColorVals['r'].indexOf('%') !== -1) ? defBultColorVals['r'].split('%').shift() : defBultColorVals['r']
    const green = (defBultColorVals['g'].indexOf('%') !== -1) ? defBultColorVals['g'].split('%').shift() : defBultColorVals['g']
    const blue = (defBultColorVals['b'].indexOf('%') !== -1) ? defBultColorVals['b'].split('%').shift() : defBultColorVals['b']
    color = toHex(255 * (Number(red) / 100)) + toHex(255 * (Number(green) / 100)) + toHex(255 * (Number(blue) / 100))
  } 
  else if (solidFill['a:prstClr']) {
    clrNode = solidFill['a:prstClr']
    const prstClr = getTextByPathList(clrNode, ['attrs', 'val'])
    color = getColorName2Hex(prstClr)
  } 
  else if (solidFill['a:hslClr']) {
    clrNode = solidFill['a:hslClr']
    const defBultColorVals = clrNode['attrs']
    const hue = Number(defBultColorVals['hue']) / 100000
    const sat = Number((defBultColorVals['sat'].indexOf('%') !== -1) ? defBultColorVals['sat'].split('%').shift() : defBultColorVals['sat']) / 100
    const lum = Number((defBultColorVals['lum'].indexOf('%') !== -1) ? defBultColorVals['lum'].split('%').shift() : defBultColorVals['lum']) / 100
    const hsl2rgb = hslToRgb(hue, sat, lum)
    color = toHex(hsl2rgb.r) + toHex(hsl2rgb.g) + toHex(hsl2rgb.b)
  } 
  else if (solidFill['a:sysClr']) {
    clrNode = solidFill['a:sysClr']
    const sysClr = getTextByPathList(clrNode, ['attrs', 'lastClr'])
    if (sysClr) color = sysClr
  }

  let isAlpha = false
  const alpha = parseInt(getTextByPathList(clrNode, ['a:alpha', 'attrs', 'val'])) / 100000
  if (!isNaN(alpha)) {
    const al_color = tinycolor(color)
    al_color.setAlpha(alpha)
    color = al_color.toHex8()
    isAlpha = true
  }

  const hueMod = parseInt(getTextByPathList(clrNode, ['a:hueMod', 'attrs', 'val'])) / 100000
  if (!isNaN(hueMod)) {
    color = applyHueMod(color, hueMod, isAlpha)
  }
  const lumMod = parseInt(getTextByPathList(clrNode, ['a:lumMod', 'attrs', 'val'])) / 100000
  if (!isNaN(lumMod)) {
    color = applyLumMod(color, lumMod, isAlpha)
  }
  const lumOff = parseInt(getTextByPathList(clrNode, ['a:lumOff', 'attrs', 'val'])) / 100000
  if (!isNaN(lumOff)) {
    color = applyLumOff(color, lumOff, isAlpha)
  }
  const satMod = parseInt(getTextByPathList(clrNode, ['a:satMod', 'attrs', 'val'])) / 100000
  if (!isNaN(satMod)) {
    color = applySatMod(color, satMod, isAlpha)
  }
  const shade = parseInt(getTextByPathList(clrNode, ['a:shade', 'attrs', 'val'])) / 100000
  if (!isNaN(shade)) {
    color = applyShade(color, shade, isAlpha)
  }
  const tint = parseInt(getTextByPathList(clrNode, ['a:tint', 'attrs', 'val'])) / 100000
  if (!isNaN(tint)) {
    color = applyTint(color, tint, isAlpha)
  }

  if (color && color.indexOf('#') === -1) color = '#' + color

  return color
}
