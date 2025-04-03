/**
 * @desc 把url设置成背景，并使其覆盖全部
 * @param newSlide ppt页对象，presentation ppt对象，url图片url
 */
function setImageToBackground(newSlide, presentation, url) {
  if (newSlide && url && presentation) {
    const image = newSlide.insertImage(url);

    // 设置图片大小和位置，使其覆盖整个幻灯片
    image.setWidth(presentation.getPageWidth());
    image.setHeight(presentation.getPageHeight());
    image.setLeft(0);
    image.setTop(0); onReplaceImage
    // 将图片移到最底层
    image.sendToBack();
  } else {
    Logger.log(newSlide, url, presentation)
  }

}
function convertPtToPx(pt) {
  return pt * 1.33; // 1 pt ≈ 1.33 px
}

function convertPxToPt(px) {
  return px * 0.75; // 1 px ≈ 0.75 pt
}
/**
 * 创建一个圆形slide.insertShape(SlidesApp.ShapeType.OVAL, position, top, Math.min(width, height), Math.min(width, height));
    
 */
function createCircleBox(slide, left, top, width, height, bold = false, title, fontSize, color, center = false) {
  const _textBox = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, convertPxToPt(left), convertPxToPt(top), convertPxToPt(Math.min(width, height)), convertPxToPt(Math.min(width, height)));
  // 让文本框的字体大小根据内容自动缩小
  const _textRange = _textBox.getText();
  // 设置目录文本
  _textRange.setText(title || " ");
  const _textStyle = _textRange.getTextStyle();
  _textStyle.setFontSize(fontSize);
  _textStyle.setBold(bold);
  _textStyle.setForegroundColor(color);
  _textStyle.setFontFamily("Open Sans");
  if (center) {
    _textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
  // 插入一个矩形形状（作为文本框）
  return _textBox;
}
/**
 * @desc 创建一个文本输入框
 */
function createTextBox(slide, left, top, width, height, bold = false, title, fontSize, color, center = false) {
  const _textBox = slide.insertTextBox("...", convertPxToPt(left), convertPxToPt(top), convertPxToPt(width), convertPxToPt(height));
  // 让文本框的字体大小根据内容自动缩小
  const _textRange = _textBox.getText();

  // 设置目录文本
  _textRange.setText((String(title) || " ").replace(/\\n/g, String.fromCharCode(10)));
  const _textStyle = _textRange.getTextStyle();
  _textStyle.setFontSize(fontSize);
  _textStyle.setBold(bold);
  _textStyle.setForegroundColor(color);
  _textStyle.setFontFamily("Georgia");
  if (center) {
    _textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
  // 插入一个矩形形状（作为文本框）
  return _textBox;
}

/**
 * 创建居中的大标题
 *  @param alignCenter boolean 是否垂直居中
 *  @return textBox TextBox
 */

function createTitleBox(slide, presentation, top, title, color, alignCenter = false) {
  const WIDTH = 400, HEIGHT = 60, FONT_SIZE = 40;
  // 宽高单位是pt
  const pageHeight = convertPtToPx(presentation.getPageHeight()), pageWidth = convertPtToPx(presentation.getPageWidth());
  const boxOffsetX = (pageWidth - WIDTH) / 2;
  const boxOffsetY = alignCenter ? (pageHeight - HEIGHT) / 2 : top;
  const _textBox = createTextBox(slide, boxOffsetX, boxOffsetY, WIDTH, HEIGHT, true, title, FONT_SIZE, color);
  const _textRange = _textBox.getText();
  _textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  return _textBox;
}
/**
 *  创建大标题2
 *  @param alignCenter boolean 是否垂直居中
 *  @return textBox TextBox
 */

function createTitleBox2(slide, presentation, title = " ", language = "en", color = "#745AE4", alignCenter = false, width = 700, top = 30) {
  const WIDTH = width, HEIGHT = 65, FONT_SIZE = 30;
  // 宽高单位是pt
  const pageHeight = convertPtToPx(presentation.getPageHeight()), pageWidth = convertPtToPx(presentation.getPageWidth());
  const boxOffsetX = (pageWidth - WIDTH) / 2;
  const boxOffsetY = alignCenter ? (pageHeight - HEIGHT) / 2 : top;
  const _textBox = createTextBox(slide, boxOffsetX, boxOffsetY, WIDTH, HEIGHT, true, title, FONT_SIZE, color);
  const _textRange = _textBox.getText();
  _textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  setAutoFit([
    {
      textBox: _textBox,
      fontSize: FONT_SIZE,
      width: WIDTH,
      rate: language === 'en' ? 0.3 : 0.7,
      height: HEIGHT,
      text: title,
      language: language
    }
  ])
  return _textBox;
}
/**
 * 插入图片并且缩放图片到指定容器大小尺寸
 */
function insertAndCropToFixedSize(offsetLeft, offsetTop, slide, imageUrl, containerWidth, containerHeight) {
  if (!imageUrl) return

  //插入图片的基础水平偏移量
  const OFFSET_LEFT = convertPxToPt(offsetLeft);
  //图片容器的基础大小
  const IMAGE_BOX_WIDTH = convertPxToPt(containerWidth), IMAGE_BOX_HEIGHT = convertPxToPt(containerHeight);
  //插入图片的基础垂直偏移量
  const OFFSET_TOP = convertPxToPt(offsetTop);
  try {
    const image = slide.insertImage(imageUrl);
    //图片实际高度与宽度
    const IMAGE_WIDTH = image.getWidth(), IMAGE_HEIGHT = image.getHeight();
    // Logger.log(`IMAGE_WIDTH${IMAGE_WIDTH},IMAGE)HEIGHT${IMAGE_HEIGHT},BOX_WIDTH:${IMAGE_BOX_WIDTH},BOX_HEIGHT:${IMAGE_BOX_HEIGHT}`)
    //如果图片是横幅，则按照宽度为基准去缩放
    if (IMAGE_WIDTH > IMAGE_HEIGHT) {
      //计算缩放比例
      let scale = 1;
      if (IMAGE_WIDTH >= IMAGE_BOX_WIDTH) {
        scale = IMAGE_BOX_WIDTH / IMAGE_WIDTH;
      } else {
        scale = IMAGE_WIDTH / IMAGE_BOX_WIDTH;
      }
      //计算缩放后的图片高度
      const imageHeight = IMAGE_HEIGHT * scale;
      //计算图片实际的垂直偏移量
      const _offsetTop = OFFSET_TOP + ((IMAGE_BOX_HEIGHT - imageHeight) / 2);
      image.setTop(_offsetTop);
      image.setLeft(OFFSET_LEFT);
      image.setWidth(IMAGE_BOX_WIDTH);
      image.setHeight(imageHeight);
    } else {
      //计算缩放比例
      let scale = 1;
      if (IMAGE_HEIGHT >= IMAGE_BOX_HEIGHT) {
        scale = IMAGE_BOX_HEIGHT / IMAGE_HEIGHT;
      } else {
        scale = IMAGE_HEIGHT / IMAGE_BOX_HEIGHT;
      }
      //计算缩放后的图片高度
      const imageWidth = IMAGE_WIDTH * scale;
      //计算图片实际的水平偏移量
      const _offsetLeft = OFFSET_LEFT + ((IMAGE_BOX_WIDTH - imageWidth) / 2);
      image.setTop(OFFSET_TOP);
      image.setHeight(IMAGE_BOX_HEIGHT);
      image.setLeft(_offsetLeft);
      image.setWidth(imageWidth);
      // return {
      //   width: imageWidth,
      //   height: IMAGE_BOX_HEIGHT,
      //   offsetLeft: offsetLeft,
      //   offsetTop: OFFSET_TOP
      // }
    }
  } catch (e) {
    //insertAndCropToFixedSizeFromImages(offsetLeft, offsetTop, slide, imageUrl, containerWidth, containerHeight)

    setDefaultImg(offsetLeft, offsetTop, slide, containerWidth, containerHeight);
    throw e;
  }
}
//移除某个编号的ppt
function removeSomePage(presentation, mark) {
  const slides = presentation.getSlides();
  try {
    const item = slides.find(it => parseInt(getOrder(it)) === parseInt(mark));
    if (item) {
      item.remove();

    }
  } catch (e) {
    throw e;
    Logger.log(e)
  }
}
/**
 * 移除Loading page
 */
function removeLoadingPageHandle(presentation, data) {
  const mark = data.order;

  removeSomePage(presentation, mark);
}
//获取模版
function getTemplate(id) {
  let _template;
  switch (id) {
    case 1:
      _template = layout1;
      break;
    case 2:
      _template = layout2;
      break;
    case 3:
      _template = layout3;
      break;
    default:
      _template = layout1;
  }
  return _template;
}
/**
 * createCoverLayout 创建封面
 * createCatalogLayout 创建目录布局
 * createChapters  创建一个章节
 * create1Title1IllustrationLayout 创建正文-> 1标题1插图(随机左右)
 * create1Title3SmallTitle1IllustrationLayout 创建正文 1大标题3小标题1图
 * create1Title3SmallTitleLayout 创建正文一大标题三小标题布局
 * create1Title4SmallTitleLayout 创建正文一大标题四小标题布局3
 * createEndingLayout 创建结尾
 * type 0结尾|1封面|2目录|3章节|4,5,6正文
 */
function useSelector(data, presentation) {
  const _template = getTemplate(data.template);
  const contents = data.content;
  if (data.construct === 1) {
    return _template.createLoadingPage;
  } else if (data.construct === 0) {
    removeLoadingPageHandle(presentation, data);
  }
  //0 1 2 3算是通用的結構，對應的也是固定的模板
  if (data.type === 0) {
    return _template.createEndingLayout
  } else if (data.type === 1) {
    return _template.createCoverLayout
  } else if (data.type === 2) {
    return _template.createCatalogLayout
  } else if (data.type === 3) {
    return _template.createChapters
  } else {
    //contentType 1、词汇库 2、语法 3、语法填空题 4、写作-填空题 5、写作-写作题 6-听力
    switch (data.contentType) {
      case 1:
        return _template.create4ItemsWith4Image;
      case 3:
        return _template.createMoreContent
      case 4:
        return _template.createMoreContent
      case 6:
        return _template.createVoicePlay;
    }
    if (data.template === 1) {
      const hasImage = (contents.some(it => Boolean(it.imageKey)) || (contents.some(it => Boolean(it.imageKeyWord)))) && (Math.random() > 0.3);
      //如果没有smallTitle
      const hasSmallTitle = contents.some(it => it.smallTitle);
      if (!hasSmallTitle) {
        return _template.createMoreContent
      }
      switch (contents.length) {
        case 1:
          return hasImage ? _template.create1Title1IllustrationLayout : _template.createMoreContent;
        case 2:
          return hasImage ? _template.create1Title2SmallTitleLayout : _template.createMoreContent;
        case 3:
          return hasImage ? _template.create1Title3SmallTitle1IllustrationLayout : Math.random() > 0.5 ? _template.create1Title3SmallTitleLayout2 : _template.create1Title3SmallTitleLayout
        case 4:
          return _template.create1Title4SmallTitleLayout
        default:
          return _template.createMoreContent
      }
    } else if (data.template === 2) {
      if (data.contentType === 10 || data.contentType === 14) {
        return _template.create3SmallTitle
      }
      const hasImage = (contents.some(it => Boolean(it.imageKey)) || (contents.some(it => Boolean(it.imageKeyWord)))) && (Math.random() > 0.3);
      const hasSmallTitle = contents.some(it => it.smallTitle);
      //沒有圖片的話不適合用那些帶圖片的模板
      if (!hasSmallTitle) {
        return _template.createMoreContent
      } else if (!hasImage) {
        return _template.createMoreContent
      }
      switch (contents.length) {
        case 1:
          return _template.create1Title1IllustrationLayout
        case 2:
          return _template.create2SamllTitle;
        case 3:
          return _template.create3SmallTitle
        case 4:
          return _template.create4SmallTitle
        default:
          return _template.createMoreContent
      }
    } else if (data.template === 3) {
      const hasImage = (contents.some(it => Boolean(it.imageKey)) || (contents.some(it => Boolean(it.imageKeyWord)))) && (Math.random() > 0.3);
      const hasSmallTitle = contents.some(it => it.smallTitle);
      //沒有圖片的話不適合用那些帶圖片的模板
      if (!hasSmallTitle) {
        return _template.createMoreContent
      } else if (!hasImage) {
        return _template.createMoreContent
      }
      switch (contents.length) {
        case 1:
          return _template.create1Title1IllustrationLayout
        case 2:
          return _template.create2SamllTitle;
        case 3:
          return _template.create3SmallTitle
        case 4:
          return _template.create4SmallTitle
        default:
          return _template.createMoreContent
      }
    }
  }
}
 
/**
 * 文本后追加超链接
 */
function appendTextWithHyperlink(textBox, text, link) {
  const textRange = textBox.getText();
  textRange.appendText("🔗"); // 追加文本

  // **必须用 TextStyle 设置超链接**
  const textStyle = textRange.getTextStyle();
  textStyle.setLinkUrl(link);
}
/**
 * 各种字体的宽度是不一样的，引入缩放系数来适应
 */
function getRate(lang) {
  const hk = 0.95, en = 0.5;
  if (lang === 'hk') return hk;
  else if (lang === 'zh') return hk;
  return en;
}
/**
 * 自動調整文字尺寸
 * @param textBoxs {textBox: Textbox, fontSize, width, height, text, rate,language}[]
 */
function setAutoFit(textBoxs = []) {
  const MIN_SIZE = 10; //最小尺寸
  const fontSizeList = [];
  textBoxs.map(it => {
    let size = it.fontSize; //原始字體尺寸
    const rate = it.rate || getRate(it.language)
    //計算每組textBox最佳的字體尺寸
    while (true) {
      if (canTextFitInTextbox(it.width, it.height, size, it.text, rate)) {
        fontSizeList.push(size);
        break;
      } else {
        if (size <= MIN_SIZE) {
          fontSizeList.push(size);
          break;
        } else {
          size -= 2;
        }
      }
    }
  })
  //選擇最小尺寸用作通用尺寸
  const minSize = Math.min(...fontSizeList) || MIN_SIZE;
  textBoxs.map(it => {
    it.textBox.getText().getTextStyle().setFontSize(convertPxToPt(minSize));
  })
  return fontSizeList;
}

/**
 * 判断textbox能否容纳文字
 */
function canTextFitInTextbox(w, h, size, text, rate = 1) {

  const charWidth = size * rate; // 假设每个字符的宽度大约是字体大小的 0.6 倍
  const charSpacing = 1; // 字距
  const lineSpacing = 2; // 行距
  const realWidth = w - 20, realHeight = h - 5;
  // 计算每行能够容纳的字符数
  const charsPerLine = Math.floor(realWidth / (charWidth + charSpacing));
  //  Logger.log("计算每行能够容纳的字符数" + charsPerLine)
  // 计算文本的总字符数
  const textLength = text.length;
  // Logger.log("计算文本的总字符数" + textLength)
  // 计算文本所需的行数
  const linesNeeded = Math.ceil(textLength / charsPerLine);
  //  Logger.log("文本所需的行数" + linesNeeded)
  // 计算每行的高度
  const lineHeight = size + lineSpacing;
  // 计算总高度
  const totalHeight = linesNeeded * lineHeight;
  //  Logger.log(totalHeight+' '+realHeight)

  // 判断是否能够容纳
  return totalHeight <= realHeight;
}
/**
 * 更换或者上传图片
 */
//更换图片
function onReplaceImage(presentationId, slideIndex, newImageUrl) {
  if (slideIndex >= 0 && newImageUrl && presentationId) {
    // 获取当前演示文稿
    const presentation = SlidesApp.openById(presentationId);
    // 获取指定的幻灯片（通过索引）
    const slide = presentation.getSlides()[slideIndex];
    // 获取幻灯片上的所有图片
    const images = slide.getImages();
    if (images.length > 0) {
      const pageWidth = presentation.getPageWidth();
      const pageHeight = presentation.getPageHeight();

      const oldImage = images.find(item => item.getWidth() < pageWidth || item.getHeight() < pageHeight);

      if (oldImage) {
        // 删除旧图片
        // oldImage.remove();
        const oldOffsetX = oldImage.getLeft(), oldOffsetY = oldImage.getTop(), oldWidth = oldImage.getWidth(), oldHeight = oldImage.getHeight();
        oldImage.remove();
        // // 插入新图片
        // const newImage = slide.insertImage(newImageUrl);
        // 缩放比例
        insertAndCropToFixedSize(oldOffsetX, oldOffsetY, slide, newImageUrl, 300, 300);

      } else {
        // 插入新图片
        const newImage = slide.insertImage(newImageUrl);
        // 缩放比例
        if (newImage) {
          const imgObj = getImagesSize(newImage, presentation);
          // 设置图片的位置和大小
          newImage.setLeft(imgObj.offsetLeft);   // 图片左边距
          newImage.setTop(imgObj.offsetTop);    // 图片上边距
          newImage.setWidth(imgObj.width);
          newImage.setHeight(imgObj.height);
        }
      }
    }
  } else {
    // 构建响应
    const response = {
      "status": "Fail",
      "message": "索引错误或者图片地址为空" + slideIndex
    };

    // 返回 JSON 响应
    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
/**
 * 删除所有幻灯片
 */
//重新生成
function removeAll(presentationId) {
  // 1. 打开 Google 幻灯片
  // const presentationId = presentationId; // 替换为你的幻灯片ID
  const presentation = SlidesApp.openById(presentationId);

  // 2. 获取所有幻灯片
  const slides = presentation.getSlides();

  // 3. 删除每一页幻灯片
  for (let i = slides.length - 1; i >= 0; i--) {
    slides[i].remove();
    Logger.log(`已删除第 ${i + 1} 页幻灯片`);
  }
}
/**
 * 返回 {width: 图片宽度, height: 图片高度, offsetLeft: 图片左偏移量, offsetTop: 图片上偏移量}
 */

function getImagesSize(image, presentation) {
  if (image && presentation) {
    // ppt 的页高和宽度
    const PAGE_WIDTH = presentation.getPageWidth(), PAGE_HEIGHT = presentation.getPageHeight();
    //插入图片的基础水平偏移量
    const OFFSET_LEFT = 410;
    //图片容器的基础大小
    const IMAGE_BOX_WIDTH = PAGE_WIDTH - OFFSET_LEFT - 40, IMAGE_BOX_HEIGHT = IMAGE_BOX_WIDTH * 3 / 4;
    //插入图片的基础垂直偏移量
    const OFFSET_TOP = (PAGE_HEIGHT - IMAGE_BOX_HEIGHT) / 2;
    //图片实际高度与宽度
    const IMAGE_WIDTH = image.getWidth(), IMAGE_HEIGHT = image.getHeight();

    //如果图片是横幅，则按照宽度为基准去缩放
    if (IMAGE_WIDTH > IMAGE_HEIGHT) {
      //计算缩放比例
      let scale = 1;
      if (IMAGE_WIDTH >= IMAGE_BOX_WIDTH) {
        scale = IMAGE_BOX_WIDTH / IMAGE_WIDTH;
      } else {
        scale = IMAGE_WIDTH / IMAGE_BOX_WIDTH;
      }
      //计算缩放后的图片高度
      const imageHeight = IMAGE_HEIGHT * scale;
      //计算图片实际的垂直偏移量
      const offsetTop = OFFSET_TOP + ((IMAGE_BOX_HEIGHT - imageHeight) / 2);
      return {
        width: IMAGE_BOX_WIDTH,
        height: imageHeight,
        offsetLeft: OFFSET_LEFT,
        offsetTop: offsetTop
      }
    } else {
      //计算缩放比例
      let scale = 1;
      if (IMAGE_HEIGHT >= IMAGE_BOX_HEIGHT) {
        scale = IMAGE_BOX_HEIGHT / IMAGE_HEIGHT;
      } else {
        scale = IMAGE_HEIGHT / IMAGE_BOX_HEIGHT;
      }
      //计算缩放后的图片高度
      const imageWidth = IMAGE_WIDTH * scale;
      //计算图片实际的水平偏移量
      const offsetLeft = OFFSET_LEFT + ((IMAGE_BOX_WIDTH - imageWidth) / 2);
      return {
        width: imageWidth,
        height: IMAGE_BOX_HEIGHT,
        offsetLeft: offsetLeft,
        offsetTop: OFFSET_TOP
      }
    }
  }
}


function generate(data = [], presentationId, title = "", template) {
  var presentationId = presentationId;  // 替换为您的演示文稿 ID
  var presentation = SlidesApp.openById(presentationId);
  const _slides = presentation.getSlides();
  const pageWidth = presentation.getPageWidth(), pageHeight = presentation.getPageHeight();
  if (title && _slides.length < 1) {
    // 使用 空白布局或其他适当的布局来设置标题
    var newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    if (template && template.title) {
      const image = newSlide.insertImage(template.title);

      // 设置图片大小和位置，使其覆盖整个幻灯片
      image.setWidth(pageWidth);
      image.setHeight(pageHeight);
      image.setLeft(0);
      image.setTop(0);
      // 将图片移到最底层
      image.sendToBack();
    }
    // 设置标题文本
    const textBox = newSlide.insertTextBox(title, 50, pageHeight / 2 - 50, pageWidth / 2 - 50, 100);
    var textRange = textBox.getText();
    // 设置标题文本
    textRange.setText(title);
    // 设置文本对齐方式为居中
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    // 设置字体样式
    var textStyle = textRange.getTextStyle();
    textStyle.setFontSize(36);
    textStyle.setBold(true);
    const regex = /\/([^\/]+)\/([^\/]+)$/;
    if (template && template.catalog) {
      const match = template.catalog.match(regex);
      if (match && match[1] === 't1') {
        textStyle.setForegroundColor('#ffffff'); // 设置文字颜色
      } else {
        textStyle.setForegroundColor('#222222'); // 设置文字颜色
      }
    } else {
      textStyle.setForegroundColor('#222222'); // 设置文字颜色
    }

  }
  if (data) {
    data.forEach(function (item, index) {
      // // 如果当前标题跟上一个标题重复则视为同一个，跳出
      // if(item.title === titleText) return
      // 使用 TITLE_AND_BODY 布局添加标题和内容
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      if (presentation.getSlides().length === 2 && template && template.catalog) {
        const image = slide.insertImage(template.catalog);
        // 设置图片大小和位置，使其覆盖整个幻灯片
        image.setWidth(pageWidth);
        image.setHeight(pageHeight);
        image.setLeft(0);
        image.setTop(0);
        // 将图片移到最底层
        image.sendToBack();
      } else if (presentation.getSlides().length > 2 && template && template.content) {
        const image = slide.insertImage(template.content);
        // 设置图片大小和位置，使其覆盖整个幻灯片
        image.setWidth(pageWidth);
        image.setHeight(pageHeight);
        image.setLeft(0);
        image.setTop(0);
        // 将图片移到最底层
        image.sendToBack();
      }
      // 获取幻灯片中的标题和内容占位符
      const placeholders = slide.getPlaceholders();
      const titlePlaceholder = placeholders[0];  // 标题占位符
      const bodyPlaceholder = placeholders[1];   // 内容占位符
      const regex = /\/([^\/]+)\/([^\/]+)$/;
      // 设置标题
      titlePlaceholder.asShape().getText().setText(item.title);
      titlePlaceholder.asShape().setLeft(40);
      // 设置内容
      bodyPlaceholder.asShape().getText().setText(item.content);
      if (template && template.catalog) {
        const match = template.catalog.match(regex);
        if (match && match[1] === 't1') {
          bodyPlaceholder.asShape().getText().getTextStyle().setForegroundColor('#ffffff');
          titlePlaceholder.asShape().getText().getTextStyle().setForegroundColor('#FFD700');
        } else {
          bodyPlaceholder.asShape().getText().getTextStyle().setForegroundColor('#111111');
          titlePlaceholder.asShape().getText().getTextStyle().setForegroundColor('#FF0000');
        }
      } else {
        bodyPlaceholder.asShape().getText().getTextStyle().setBaselineOffset
        bodyPlaceholder.asShape().getText().getTextStyle().setForegroundColor('#111111');
        titlePlaceholder.asShape().getText().getTextStyle().setForegroundColor('#FF0000');
      }
      //设置行间距 150
      // bodyPlaceholder.asShape().getText().getTextStyle().setLineSpacing(150)
      const paragraphs = bodyPlaceholder.asShape().getText()?.getParagraphs?.();
      paragraphs?.forEach?.((paragraph) => {
        const style = paragraph.getRange().getParagraphStyle();
        style.setLineSpacing(150); // 设置为 1.5 倍行距（150%）
      });
      bodyPlaceholder.asShape().getText().getTextStyle().setFontSize(14);
      //.setFontSize(24)
      bodyPlaceholder.asShape().setWidth(360);
      bodyPlaceholder.asShape().setLeft(40);
      // bodyPlaceholder.asShape().setLeft(20);
      // bodyPlaceholder.asShape().setTop(20);
      // 设置标题和内容的样式
      // titlePlaceholder.asShape().getText().setFontSize(36).setBold(true);
      // bodyPlaceholder.asShape().getText().setFontSize(24);
      // 添加图片到幻灯片
      if (item.image) {
        var image = slide.insertImage(item.image);
        const imgObj = getImagesSize(image, presentation);

        // 设置图片的位置和大小
        image.setLeft(imgObj.offsetLeft);   // 图片左边距
        image.setTop(imgObj.offsetTop);    // 图片上边距


        // 设置图片的新大小
        image.setWidth(imgObj.width);
        image.setHeight(imgObj.height);
      }

    })
  }
}
/**
 * 删除最后一张幻灯片
 */
function removeTheLastOne(presentationId) {
  var presentationId = presentationId;  // 替换为您的演示文稿 ID
  var presentation = SlidesApp.openById(presentationId);
  const slides = presentation.getSlides();
  if (slides && slides.length > 0) {
    const theLastSlide = slides.pop();
    theLastSlide && theLastSlide.remove();
  }
}

//更换背景图片
function changeBackground(presentationId) {
  var presentation = SlidesApp.openById(presentationId);
  const slides = presentation.getSlides();
  // 3. 从 Google Drive 获取图片的文件 ID
  const imageFileId = '1hU8EUKSI9ysSNbsjnnTVTvfs_2n7Hnsp'; // 替换为你的图片文件ID
  const imageFile = DriveApp.getFileById(imageFileId);

  // 4. 获取图片的 Blob 对象
  const imageBlob = imageFile.getBlob();
  // 5. 获取幻灯片的页面尺寸
  const pageWidth = presentation.getPageWidth();
  const pageHeight = presentation.getPageHeight();

  // 6. 遍历每张幻灯片并插入图片
  for (let slide of slides) {
    // 插入图片
    const image = slide.insertImage(imageBlob);

    // 设置图片大小和位置，使其覆盖整个幻灯片
    image.setWidth(pageWidth);
    image.setHeight(pageHeight);
    image.setLeft(0);
    image.setTop(0);

    // 将图片移到最底层
    image.sendToBack();
  }
  const response = {
    "status": "Success",
    "message": "success"
  };
  // 返回 JSON 响应
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function getGoogleApi(text) {
  return `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(text)}&cx=546376535b2624a24&key=AIzaSyAGpr7NAyd-2C2iJGhp43ieXj45SC4ZStc&searchType=image&rights=cc_publicdomain`
}

function searchGoogleImages(text) {
  const api = getGoogleApi(text);
  const response = UrlFetchApp.fetch(api, { muteHttpExceptions: true });
  if (response.getResponseCode() === 200) {
    try {
      const data = JSON.parse(response.getContentText())
      if (data?.items instanceof Array) {
        const result = data.items.map(it => it.link);
        return result
      }
      return []
    } catch (e) {
      Logger.log(e)
      throw e;
      return []
    }

  }
}
/**
 * 插入占位 loading gif
 */
function insertPlaceholderGif(offsetLeft, offsetTop, slide, containerWidth, containerHeight) {

  const file = DriveApp.getFileById(LOADING_GIF_FILE_ID);
  const blob = file.getBlob();
  //插入图片的基础水平偏移量
  const OFFSET_LEFT = convertPxToPt(offsetLeft);
  //图片容器的基础大小
  const IMAGE_BOX_WIDTH = convertPxToPt(containerWidth), IMAGE_BOX_HEIGHT = convertPxToPt(containerHeight);
  //插入图片的基础垂直偏移量
  const OFFSET_TOP = convertPxToPt(offsetTop);
  const image = slide.insertImage(blob);
  //图片实际高度与宽度
  const IMAGE_WIDTH = image.getWidth(), IMAGE_HEIGHT = image.getHeight();
  // Logger.log(`IMAGE_WIDTH${IMAGE_WIDTH},IMAGE)HEIGHT${IMAGE_HEIGHT},BOX_WIDTH:${IMAGE_BOX_WIDTH},BOX_HEIGHT:${IMAGE_BOX_HEIGHT}`)
  //如果图片是横幅，则按照宽度为基准去缩放
  if (IMAGE_WIDTH > IMAGE_HEIGHT) {
    //计算缩放比例
    let scale = 1;
    if (IMAGE_WIDTH >= IMAGE_BOX_WIDTH) {
      scale = IMAGE_BOX_WIDTH / IMAGE_WIDTH;
    } else {
      scale = IMAGE_WIDTH / IMAGE_BOX_WIDTH;
    }
    //计算缩放后的图片高度
    const imageHeight = IMAGE_HEIGHT * scale;
    //计算图片实际的垂直偏移量
    const _offsetTop = OFFSET_TOP + ((IMAGE_BOX_HEIGHT - imageHeight) / 2);
    image.setTop(_offsetTop);
    image.setLeft(OFFSET_LEFT);
    image.setWidth(IMAGE_BOX_WIDTH);
    image.setHeight(imageHeight);
  } else {
    //计算缩放比例
    let scale = 1;
    if (IMAGE_HEIGHT >= IMAGE_BOX_HEIGHT) {
      scale = IMAGE_BOX_HEIGHT / IMAGE_HEIGHT;
    } else {
      scale = IMAGE_HEIGHT / IMAGE_BOX_HEIGHT;
    }
    //计算缩放后的图片高度
    const imageWidth = IMAGE_WIDTH * scale;
    //计算图片实际的水平偏移量
    const _offsetLeft = OFFSET_LEFT + ((IMAGE_BOX_WIDTH - imageWidth) / 2);
    image.setTop(OFFSET_TOP);
    image.setHeight(IMAGE_BOX_HEIGHT);
    image.setLeft(_offsetLeft);
    image.setWidth(imageWidth);
    // return {
    //   width: imageWidth,
    //   height: IMAGE_BOX_HEIGHT,
    //   offsetLeft: offsetLeft,
    //   offsetTop: OFFSET_TOP
    // }
  }
  return image;
}
/**
 * 从网络加载图片
 */
function fetchImage(url, goOn = false) {
  if (url) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        const blob = response.getBlob();
        return blob;
      } else if (goOn) {
        return false
      }
    } catch (e) {
      throw e;

    }
  }
  const file = DriveApp.getFileById(DEFAULT_IMAGE_FILE_ID);
  return file.getBlob();
}

/**
 * 从列表里得到一份可用的图片blob
 */
function getRealBlob(urls, count = 0) {
  const blob = fetchImage(urls[count], urls.length > 0);
  if (!blob && count < urls.length) {
    getRealBlob(urls, count + 1);
  } else if (blob) {
    return blob;
  } else {
    const file = DriveApp.getFileById(DEFAULT_IMAGE_FILE_ID);
    return file.getBlob();
  }
}
function setDefaultImg(offsetLeft, offsetTop, slide, containerWidth, containerHeight) {
  const file = DriveApp.getFileById(DEFAULT_IMAGE_FILE_ID);
  const blob = file.getBlob();
  insertAndCropToFixedSize(offsetLeft, offsetTop, slide, blob, containerWidth, containerHeight)
}
/**
 * 多张图片轮番重试
 */
function insertAndCropToFixedSizeFromImages(offsetLeft, offsetTop, slide, imageUrl, containerWidth, containerHeight, reverse = false, count = 0) {
  if (!imageUrl) return

  try {
    const blob = getRealBlob(imageUrl, count);
    //插入图片的基础水平偏移量
    const OFFSET_LEFT = convertPxToPt(offsetLeft);
    //图片容器的基础大小
    const IMAGE_BOX_WIDTH = convertPxToPt(containerWidth), IMAGE_BOX_HEIGHT = convertPxToPt(containerHeight);
    //插入图片的基础垂直偏移量
    const OFFSET_TOP = convertPxToPt(offsetTop);
    const image = slide.insertImage(blob);
    if (reverse) {
      image.setWidth(containerWidth);
      image.setHeight(containerHeight);
      image.setTop(convertPxToPt(offsetTop));
      image.setLeft(convertPxToPt(offsetLeft));
      return image
    }
    //图片实际高度与宽度
    const IMAGE_WIDTH = image.getWidth(), IMAGE_HEIGHT = image.getHeight();

    // Logger.log(`IMAGE_WIDTH${IMAGE_WIDTH},IMAGE)HEIGHT${IMAGE_HEIGHT},BOX_WIDTH:${IMAGE_BOX_WIDTH},BOX_HEIGHT:${IMAGE_BOX_HEIGHT}`)
    //如果图片是横幅，则按照宽度为基准去缩放
    if (IMAGE_WIDTH > IMAGE_HEIGHT) {
      //计算缩放比例
      let scale = 1;
      if (IMAGE_WIDTH >= IMAGE_BOX_WIDTH) {
        scale = IMAGE_BOX_WIDTH / IMAGE_WIDTH;
      } else {
        scale = IMAGE_WIDTH / IMAGE_BOX_WIDTH;
      }
      //计算缩放后的图片高度
      const imageHeight = IMAGE_HEIGHT * scale;
      //计算图片实际的垂直偏移量
      const _offsetTop = OFFSET_TOP + ((IMAGE_BOX_HEIGHT - imageHeight) / 2);
      image.setTop(_offsetTop);
      image.setLeft(OFFSET_LEFT);
      image.setWidth(IMAGE_BOX_WIDTH);
      image.setHeight(imageHeight);
    } else {
      //计算缩放比例
      let scale = 1;
      if (IMAGE_HEIGHT >= IMAGE_BOX_HEIGHT) {
        scale = IMAGE_BOX_HEIGHT / IMAGE_HEIGHT;
      } else {
        scale = IMAGE_HEIGHT / IMAGE_BOX_HEIGHT;
      }
      //计算缩放后的图片高度
      const imageWidth = IMAGE_WIDTH * scale;
      //计算图片实际的水平偏移量
      const _offsetLeft = OFFSET_LEFT + ((IMAGE_BOX_WIDTH - imageWidth) / 2);
      image.setTop(OFFSET_TOP);
      image.setHeight(IMAGE_BOX_HEIGHT);
      image.setLeft(_offsetLeft);
      image.setWidth(imageWidth);

    }
    reverse && image.setHeight(containerHeight);

  } catch (e) {
    if (imageUrl[count + 1]) {
      insertAndCropToFixedSizeFromImages(offsetLeft, offsetTop, slide, imageUrl, containerWidth, containerHeight, reverse, count + 1);
    } else {
      setDefaultImg(offsetLeft, offsetTop, slide, containerWidth, containerHeight)
    }
    Logger.log(e)
    throw e;
  }
}

/**
 * 设置备注（slide编号）
 */
function setOrder(slide, text) {
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  notesShape.getText().setText(text || "100");
}

/**
 * 拿到备注(slide编号)
 */
function getOrder(slide) {
  const notesShape = slide.getNotesPage().getSpeakerNotesShape();
  const text = notesShape.getText().asString();
  return text;
}
/**
 * 整个ppt重新排序
 */
function setSequence(presentation) {
  const slides = presentation.getSlides();
  for (let i = 0; i < slides.length; i++) {
    const _slides = presentation.getSlides();
    const index = presentation.getSlides().findIndex(it => parseInt(getOrder(it)) - 1 === i);
    if (index > -1) {
      _slides[index].move(i);
    }
  }
}
function insertVoice(textBox, audioFileId) {
  const audioUrl = "https://drive.google.com/file/d/" + audioFileId + "/preview"; // 直接播放链接

  // 设置音频链接（不会跳转到新页面）
  textBox.getText().getTextStyle().setLinkUrl(audioUrl);
}
function test() {
  searchGoogleImages("iron man")
}