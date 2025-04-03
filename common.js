/**
 * 在幻灯片指定位置插入自动调整字体大小的文本框（支持微软雅黑）
 * @param {Slide} slide - 目标幻灯片对象
 * @param {string} text - 要插入的文本（支持用 \n 换行）
 * @param {number} x - 文本框X坐标
 * @param {number} y - 文本框Y坐标
 * @param {number} textBoxWidthPt - 文本框宽度（磅）
 * @param {number} textBoxHeightPt - 文本框高度（磅）
 * @param {string} [fontFamily="Microsoft YaHei"] - 字体类型（默认微软雅黑）
 * @param {number} [initialFontSizePt=72] - 字体大小（默认24）
 * @param {number} [color="#ffffff"] - 默认字体颜色（默认白色）
 * @param {number} [hAlign=false] - 水平对齐方式（默认false）
 * @param {number} [vAlign=false] - 垂直对齐方式（默认false）
 * @param {number} [isBold=false] - 是否加粗（默认false）
 */
function insertAutoFitText(slide, textContent=" ", x, y, textBoxWidthPt, textBoxHeightPt, initialFontSizePt=24, color="#ffffff", hAlign=false, vAlign=false, isBold=false,fontFamily= "Georgia", ) {
  // 预检测初始字体是否适用
  const needAdjustment = !checkFontSizeFits(
    textContent,
    textBoxWidthPt,
    textBoxHeightPt,
    initialFontSizePt,
    fontFamily,
    isBold
  );

  // 根据检测结果选择字体大小
  let result;
  if (needAdjustment) {
    result = calculateOptimalFontSize(
      textContent,
      textBoxWidthPt,
      textBoxHeightPt-8,
      initialFontSizePt,
      fontFamily,
      isBold
    );
  } else {
    result = {
      fontSize: initialFontSizePt,
      lines: calculateLines(
        textContent,
        textBoxWidthPt,
        initialFontSizePt-8,
        fontFamily,
        isBold
      )
    };
  }

  // 插入文本框并设置样式
  const textBox = slide.insertTextBox(
    textContent,
    x, y,          // 坐标位置（单位：pt）
    textBoxWidthPt,
    textBoxHeightPt
  );
  
  const textStyle = textBox.getText().getTextStyle();
  textStyle
    .setFontFamily(fontFamily)
    .setBold(isBold)
    .setFontSize(result.fontSize)
    .setForegroundColor(color);

  // 设置对齐方式
  setTextAlignment(textBox, hAlign, vAlign, result, textBoxHeightPt)

  return textBox  
}
// 字体尺寸适用性检测
function checkFontSizeFits(text, widthPt, heightPt, fontSize, fontName, isBold) {
  const LINE_HEIGHT_FACTOR = 1.2;
  const lines = calculateLines(text, widthPt, fontSize, fontName, isBold);
  Logger.log("lines==>"+lines)
  Logger.log("fontSize==>"+fontSize)
  Logger.log("heightPt==>"+heightPt)
  Logger.log("LINE_HEIGHT_FACTOR==>"+(lines * fontSize * LINE_HEIGHT_FACTOR))
  Logger.log("FLAG==>"+(lines * fontSize * LINE_HEIGHT_FACTOR <= (heightPt-20)))
  return lines * fontSize * LINE_HEIGHT_FACTOR <= (heightPt-8);
}
// 独立行数计算器
function calculateLines(text, widthPt, fontSize, fontName, isBold) {
  const FONT_METRICS = {
    '微软雅黑': { ch: 1.0, en: 0.65 },
    'Arial': { ch: 1.0, en: 0.60 },
    'default': { ch: 1.0, en: 0.60 },
    'Georgia': { ch: 1.0, en: 0.60}
  };
  const BOLD_FACTOR = isBold ? 1.1 : 1;
  const metric = FONT_METRICS[fontName] || FONT_METRICS.default;

  let lines = 0;
  const paragraphs = text.split('\n');
  
  paragraphs.forEach(para => {
    let lineWidth = 0;
    lines++; // 每个段落至少一行
    
    for (const char of para) {
      const isChinese = /[\u4E00-\u9FFF]/.test(char);
      const charWidth = fontSize * BOLD_FACTOR * (isChinese ? metric.ch : metric.en);
      
      if (lineWidth + charWidth > widthPt) {
        lines++;
        lineWidth = charWidth;
      } else {
        lineWidth += charWidth;
      }
    }
  });
  
  return lines;
}
function calculateOptimalFontSize(text, widthPt, heightPt, maxSize, fontName, isBold) {
  const FONT_METRICS = {
    '微软雅黑': { ch: 1.0, en: 0.65 },
    'Arial': { ch: 1.0, en: 0.60 },
    'default': { ch: 1.0, en: 0.60 }
  };
  const LINE_HEIGHT_FACTOR = 1.2;
  const BOLD_FACTOR = isBold ? 1.1 : 1;
  const MIN_FONT_SIZE = 4;

  let bestSize = MIN_FONT_SIZE;
  let bestLines = 1;
  const metric = FONT_METRICS[fontName] || FONT_METRICS.default;

  let low = MIN_FONT_SIZE;
  let high = maxSize;

  while (high - low > 0.1) {
    const mid = (low + high) / 2;
    let lineWidth = 0;
    let lines = 1;

    // 分段处理换行符
    const paragraphs = text.split('\n');
    for (const para of paragraphs) {
      let currentLineWidth = 0;
      for (const char of para) {
        const isChinese = /[\u4E00-\u9FFF]/.test(char);
        const charWidth = mid * BOLD_FACTOR * (isChinese ? metric.ch : metric.en);
        
        if (currentLineWidth + charWidth > widthPt) {
          lines++;
          currentLineWidth = charWidth;
        } else {
          currentLineWidth += charWidth;
        }
      }
      lines++; // 换行符自动增加行数
    }
    lines--; // 修正最后一个换行符

    const totalHeight = lines * mid * LINE_HEIGHT_FACTOR;
    
    if (totalHeight <= heightPt) {
      bestSize = mid;
      bestLines = lines;
      low = mid;
    } else {
      high = mid;
    }
  }
  return {
    fontSize: Number(bestSize.toFixed(1)),
    lines: bestLines
  };
}
//设置对齐方式
function setTextAlignment(textBox, hAlign, vAlign, calcResult, containerHeight) {
  let horizontalAlignment = SlidesApp.ParagraphAlignment.CENTER; // START/CENTER/END
  let verticalAlignment = 'MIDDLE'; // TOP/MIDDLE/BOTTOM

  const LINE_HEIGHT_FACTOR = 1.2;
  
  // 设置水平对齐
  if(hAlign){
    textBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
  if(vAlign){
    // 设置垂直对齐
    const paragraphs = textBox.getText().getParagraphs();
    if (paragraphs.length === 0) return;

    // 计算垂直间距
    const totalHeight = calcResult.lines * calcResult.fontSize * LINE_HEIGHT_FACTOR;
    let spaceAbove = 0;
    
    switch (verticalAlignment.toUpperCase()) {
      case 'MIDDLE':
        spaceAbove = (containerHeight - totalHeight) / 2;
        break;
      case 'BOTTOM':
        spaceAbove = containerHeight - totalHeight;
        break;
    }

    // 设置首行间距
    const firstParagraphRange = paragraphs[0].getRange();
    firstParagraphRange.getParagraphStyle().setSpaceAbove(Math.max(spaceAbove, 0));

    // 设置统一行高
    const paragraphStyle = textBox.getText().getParagraphStyle();
    paragraphStyle.setLineSpacing(LINE_HEIGHT_FACTOR * 100);
  }

}

//设置同样布局的文本框的文本字体统一大小，采用其中最小的作为合适的字体
function setMinFontSize(textBoxs = []) {
  if(textBoxs.length>0){
    let MIN_SIZE = textBoxs[0].getText().getTextStyle().getFontSize(); //最小尺寸
    for (const v of textBoxs) {
      let cur = v.getText().getTextStyle().getFontSize();
      if (cur < MIN_SIZE) {
        MIN_SIZE = cur
      }
    }
    textBoxs.map(item => {
      item.getText().getTextStyle().setFontSize(MIN_SIZE);
    })
    return MIN_SIZE;
  }
}