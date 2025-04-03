
var layout5 = {
  /**
   * 创建模板5-1 封面
   */
  createCoverLayout(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE5[0].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    //创建标题
    insertAutoFitText(
      newSlide,
      objData.title||" ",
      180, 100, 400, 200, 40, 
      "#3BA9D9",
      true,
      true
    )
  },

  /**
   * 创建模板5-2 目录
   */
  // createCatalogLayout(newSlide, presentation, objData={}){
  //   //设置背景图片
  //   const fileId = TEMPLATE5[1].url;
  //   const file = DriveApp.getFileById(fileId);
  //   var blob = file.getBlob();
  //   setImageToBackground(newSlide, presentation, blob);
  //   //创建标题
  //   const pageWidth = presentation.getPageWidth();
  //   const titleBox = insertAutoFitText(
  //     newSlide,
  //     "目录",
  //     (pageWidth - 100) / 2, 40, 100, 40, 30, 
  //     "#3BA9D9",
  //     true,
  //     true
  //   )
  //   for (let i = 0; i < objData.content.length; i++) {
  //     const C_TOP = 100, C_LEFT = 100
  //     if(i%2 === 0){
  //       const _top = C_TOP+(i/2*50)
  //       insertNumberedCircle(newSlide,i+1,C_LEFT, _top)
  //       insertAutoFitText(
  //         newSlide,
  //         objData.content[i].smallTitle||" ",
  //         C_LEFT+40, _top, 200, 40, 18, 
  //         "#404040",
  //       )
  //     }else{
  //       const _top = C_TOP+((i-1)/2*50)
  //       const _left = C_LEFT + 270
  //       insertNumberedCircle(newSlide,i+1,_left, _top)
        
  //       insertAutoFitText(
  //         newSlide,
  //         objData.content[i].smallTitle||" ",
  //         C_LEFT+310, _top, 200, 40, 18, 
  //         "#404040",
  //       )

  //     }
  //   }
  // },
  createCatalogLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE5[1].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    insertAutoFitText(
      newSlide,
      objData.language === 'hk' ? "目錄" : objData.language === "zh" ? "目录" : "Contents",
      (pageWidth - 100) / 2, 40, 100, 40, 30, 
      "#000000",
      false,
      false
    )
    // createTitleBox(newSlide, presentation, 20, objData.language === 'hk' ? "目錄" : objData.language === "zh" ? "目录" : "Contents", "#E32062", alignCenter = false)
    const _items = [], { length } = objData.content;
    if (length < 1) return;
    const LEFT = 205, TOP = length <= 5 ? 75 : length > 6 && length < 9 ? 110 : length >= 9 ? 85 : 140, HEIGHT = 40, WIDTH = length > 5 ? 300 : 300 * 2, FONT_SIZE = 26;
    const range = length < 5 ? 5 : length === 6 ? 3 : length === 7 ? 4 : length === 8 ? 4 : 5;
    for (let i = 0; i < length; i++) {
      if (i < range) {
        //创建1-5数字文本框
        const _shape = createCircleBox(newSlide, 155, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)) + 60, 44, 44, true, i + 1, 15, "#ffffff").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        _shape.getFill().setSolidFill("#3BA9D9");
        _shape.getBorder().setTransparent();

        const _textBox = createTextBox(newSlide, LEFT, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)) + 60, WIDTH, HEIGHT, false, objData.content?.[i]?.smallTitle || " ", FONT_SIZE, "#404040");
        _items.push({
          textBox: _textBox,
          fontSize: FONT_SIZE,
          width: WIDTH,
          height: HEIGHT,
          rate: objData.language === 'en' ? 0.3 : 0.7,
          text: objData.content[i].smallTitle,
          language: objData.language
        })
        _textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      } else {
        const _shape = createCircleBox(newSlide, 175 + WIDTH + 50, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + ((i - range) * (HEIGHT + 25)) + 60, 44, 44, i < 9, i + 1, i < 9 ? 15 : 7, "#ffffff").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        _shape.getFill().setSolidFill("#3BA9D9");
        _shape.getBorder().setTransparent();

        const _textBox = createTextBox(newSlide, 575, TOP + (length > (i - range) ? 0 : ((HEIGHT + 25) * ((i - range) - length) / 2)) + ((i - range) * (HEIGHT + 25)) + 60, WIDTH, HEIGHT, false, objData.content[i].smallTitle || " ", FONT_SIZE, "#404040");
        _items.push({
          textBox: _textBox,
          fontSize: FONT_SIZE,
          width: WIDTH,
          rate: objData.language === 'en' ? 0.3 : 0.7,
          height: HEIGHT,
          text: objData.content[i].smallTitle,
          language: objData.language
        })
        _textBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
      }
    }
    setAutoFit(_items);

  },

  /**
   * 创建模板5-3 章节
   */
  createChapters(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE5[2].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    const pageHeight = presentation.getPageHeight();
    insertNumberedCircle(newSlide,objData.chapters,180, (pageHeight - 60) / 2, 34,60)

    insertAutoFitText(
      newSlide,
      objData.title||" ",
      180, 100, 400, 200, 40, 
      "#3BA9D9",
      true,
      true
    )
  },

  /**
   * 创建正文-5-4-1> 1标题1插图(随机左右)
   */
  create1Title1IllustrationLayout(newSlide, presentation, objData) {
    //设置背景图片
    const fileId = TEMPLATE5[4].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    //创建标题
    insertAutoFitText(
      newSlide,
      objData.title||" ",
      (pageWidth - 300) / 2, 40, 300, 50, 22, 
      "#3BA9D9",
      true,
      false
    )
    
    const imageUrl = objData.content.find(it => Boolean(it.imageKey))?.imageKey; //"https://cdn.pixabay.com/photo/2024/12/27/14/58/owl-9294302_640.jpg"
    const __keyword = objData.content?.find(it => it.imageKeyWord)?.imageKeyWord;

    //创建图右文左布局
    const createLeftLayout = function () {
      //创建标题
      insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle,
        70, 100, 300, 50, 22, 
        "#3BA9D9",
        false,
        false
      )

      //创建内容
      const _textBox = insertAutoFitText(
        newSlide,
        objData.content[0].smallContent,
        70, 150, 280, 200, 16, 
        "#595959",
        false,
        false
      )
     

      //创建链接
      if (objData.content[0].link) {
        appendTextWithHyperlink(_textBox, objData.content[0].text, objData.content[0].link)
      }

      //创建图片
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(500, 130, newSlide, 350, 350);
        const urls = searchGoogleImages(__keyword);

        insertAndCropToFixedSizeFromImages(500, 130, newSlide, urls, 350, 350);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(500, 130, newSlide, imageUrl, 350, 350);
      }
    }
    //创建图左文右布局
    const createRightayout = function () {
      //创建标题
      insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle,
        350, 100, 300, 50, 20, 
        "#595959",
        false,
        false
      )

      //创建内容
      const _textBox = insertAutoFitText(
        newSlide,
        objData.content[0].smallContent,
        350, 150, 280, 200, 14, 
        "#595959",
        false,
        false
      )
     

      //创建链接
      if (objData.content.link) {
        appendTextWithHyperlink(_textBox, objData.content[0].text, objData.content[0].link)
      }

      //创建图片
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(100, 130, newSlide, 350, 350);
        const urls = searchGoogleImages(__keyword);

        insertAndCropToFixedSizeFromImages(100, 130, newSlide, urls, 350, 350);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(100, 130, newSlide, imageUrl, 350, 350);
      }
    }
    Math.random() > 0.5 ? createRightayout() : createLeftLayout();
    
  },

  /**
   * 创建模板5-4-2 两个小标题
   */
  create2SamllTitle(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE5[4].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    //创建标题
    insertAutoFitText(
      newSlide,
      objData.title||" ",
      (pageWidth - 300) / 2, 40, 300, 50, 22, 
      "#3BA9D9",
      true,
      false
    )

  
    let textBox1 = [], textBox2 = []  
    if(objData.content && objData.content[0]){
      const  text1= insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle||" ",
        90, 140, 150, 30, 14, 
        "#000000",
        false,
        true
      )
      const  text2= insertAutoFitText(
        newSlide,
        objData.content[0].smallContent||" ",
        90, 180, 370, 50, 12, 
        "#595959",
        false,
        false
      )
      textBox1.push(text1)
      textBox2.push(text2)
      if (objData.content[0]?.imageKey) {
        insertAndCropToFixedSize(640, 220, newSlide, objData.content[0]?.imageKey, 120, 100);
      }
    }

    // if(objData.content && objData.content[1]){
    //   const text1 = insertAutoFitText(
    //     newSlide,
    //     objData.content[1].smallTitle||" ",
    //     90, 250, 150, 30, 14, 
    //     "#000000",
    //     false,
    //     true
    //   )
    //   const text2 = insertAutoFitText(
    //     newSlide,
    //     objData.content[1].smallContent||" ",
    //     90, 290, 370, 50, 12, 
    //     "#000000",
    //     false,
    //     false
    //   )
    //   textBox1.push(text1)
    //   textBox2.push(text2)

    //   if (objData.content[1]?.imageKey) {
    //     insertAndCropToFixedSize(640, 360, newSlide, objData.content[01]?.imageKey, 120, 100);
    //   }
    // }
    setMinFontSize(textBox1)
    setMinFontSize(textBox2)
  },
  

  /**
   * 创建模板5-10 结尾
   */
  createEndingLayout(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE5[0].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    //创建标题
    insertAutoFitText(
      newSlide,
      objData.title||" ",
      180, 100, 400, 200, 40, 
      "#3BA9D9",
      true,
      true
    )
    
  },
}

/**
 * 插入一个带圆形背景的序号
 */
function insertNumberedCircle(slide, number, x, y, fontSize = 16, diameter = 35) {
  // 配置参数
  const circleConfig = {
    number: number,               // 显示的数字
    position: { x: x, y: y }, // 位置（单位：pt）
    diameter: diameter,           // 直径（单位：pt）
    bgColor: '#3BA9D9',      // 背景颜色（Google蓝色）
    textColor: '#FFFFFF',    // 文字颜色
    fontSize: fontSize,            // 字体大小
    fontFamily: 'Arial'      // 字体类型
  };

  // 创建圆形
  const circle = slide.insertShape(
    SlidesApp.ShapeType.ELLIPSE,
    circleConfig.position.x,
    circleConfig.position.y,
    circleConfig.diameter,
    circleConfig.diameter
  );

  // 设置圆形样式
  const fill = circle.getFill();
  fill.setSolidFill(circleConfig.bgColor);
  circle.getBorder().setTransparent(); // 移除边框

  insertAutoFitText(
    slide,
    number.toString(),
    x, y, diameter, diameter, fontSize, 
    "#FFFFFF",
    true
  )

  // 添加数字文本
  // const textRange = circle.getText().setText(circleConfig.number);
  
  // // 设置文本样式
  // const textStyle = textRange.getTextStyle();
  // textStyle
  //   .setForegroundColor(circleConfig.textColor)
  //   .setFontSize(circleConfig.fontSize)
  //   .setFontFamily(circleConfig.fontFamily)
  //   .setBold(false);

  // // 设置文本居中
  // const paragraphStyle = textRange.getParagraphStyle();
  // paragraphStyle
  //   .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER)    // 水平居中
  //   .setLineSpacing(100) // 100%行高保证垂直居中
  //   .setSpaceAbove(0);   // 移除额外间距
}

function createCustomCapsule(slide) {
  const capsuleWidth = 200;
  const capsuleHeight = 60;
  const radius = capsuleHeight / 2; // 半圆半径

  // 1. 创建左半圆
  const leftCircle = slide.insertShape(
    SlidesApp.ShapeType.ELLIPSE,
    100,           // X
    100,           // Y
    radius * 2,    // 直径宽度
    capsuleHeight  // 高度
  );

  // 2. 创建中间矩形
  const rect = slide.insertShape(
    SlidesApp.ShapeType.RECTANGLE,
    100 + radius,  // X
    100,           // Y
    capsuleWidth - radius * 2, // 宽度
    capsuleHeight  // 高度
  );

  // 3. 创建右半圆
  const rightCircle = slide.insertShape(
    SlidesApp.ShapeType.ELLIPSE,
    100 + capsuleWidth - radius * 2, // X
    100,                             // Y
    radius * 2,                      // 宽度
    capsuleHeight                    // 高度
  );

  // 4. 组合形状并设置统一样式
  const shapes = [leftCircle, rect, rightCircle];
  shapes.forEach(shape => {
    shape.getFill().setSolidFill("#4285F4");
    shape.getBorder().setTransparent();
  });

  // 5. 添加文本（需单独创建文本框）
  const textBox = slide.insertTextBox(
    "按钮文字",
    100 + radius,  // X
    100 + 10,      // Y（垂直居中偏移）
    capsuleWidth - radius * 2,
    capsuleHeight - 20
  );
  textBox.getText().getTextStyle()
    .setFontSize(18)
    .setForegroundColor("#FFFFFF")
    // .setHorizontalAlignment(SlidesApp.HorizontalAlignment.CENTER);
}


function _layou5Test() {
  const ID = '1uV_BHGuGHgdFybDZE7tXiBsokJl1tiU5Fdmad9KWYaw';
  const presentation = SlidesApp.openById(ID);
  const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  const objData = {
      "template": 3,
      "pptId": "1nFOVQ000Gh6VctLdhBXvaGIaQ5zlCJTOkeTTLmzmlMA",
      "chapters": 2,
      "language": "en",
      "construct": 1,
      "title": "Reading comprehension",
      "type": 7,
      "contentType": 0,
      "pptType": 1,
      "content": [
          {
              "smallTitle": "Understanding Historical Contexts",
              "imageKeyWord": "Hong Kong historical events",
              "smallContent": "In this reading comprehension section, examine the historical contexts behind key events in Hong Kong's history. Analyze texts to extract meaningful insights and improve your understanding of how historical narratives shape the present and future of Hong Kong."
          },
          {
              "smallTitle": "Analyzing Cultural Transformations",
              "imageKeyWord": "Hong Kong cultural transformation",
              "smallContent": "Explore the cultural transformations in Hong Kong's history. This part of the reading comprehension exercises will guide you through significant cultural and societal changes that have influenced the city\u2019s identity."
          }
      ],
      "order": 2
    }
  layout5.create2SamllTitle(newSlide,presentation,objData)

}