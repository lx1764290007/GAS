
var layout3 = {
  /**
   * 创建模板3-1  封面
   */
  createCoverLayout(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[0].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    //创建标题
    insertAutoFitText(
      newSlide,
      objData.title,
      50, 100, 300, 150, 34, 
      "#ffffff",
      true,
      true
    )
    
  },

  /**
   * 创建模板3-2 目录
   */
  createCatalogLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE3[6].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    insertAutoFitText(
      newSlide,
      objData.language === 'hk' ? "目錄" : objData.language === "zh" ? "目录" : "Contents",
      80, 20, 250, 60, 22, 
      "#ffffff",
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
        const _shape = createCircleBox(newSlide, 155, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)) + 60, 44, 44, true, i + 1, 15, "#000000").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        // _shape.getFill().setSolidFill("#ffffff");
        _shape.getBorder().setTransparent();
        // _shape.setShapeType(SlidesApp.ShapeType.OVAL);

        const _textBox = createTextBox(newSlide, LEFT, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)) + 60, WIDTH, HEIGHT, false, objData.content?.[i]?.smallTitle || " ", FONT_SIZE, "#000000");
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
        const _shape = createCircleBox(newSlide, 175 + WIDTH + 50, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + ((i - range) * (HEIGHT + 25)) + 60, 44, 44, i < 9, i + 1, i < 9 ? 15 : 7, "#000000").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        // _shape.getFill().setSolidFill("#ffffff");
        _shape.getBorder().setTransparent();

        //  _shape.setShapeType(SlidesApp.ShapeType.OVAL);
        const _textBox = createTextBox(newSlide, 575, TOP + (length > (i - range) ? 0 : ((HEIGHT + 25) * ((i - range) - length) / 2)) + ((i - range) * (HEIGHT + 25)) + 60, WIDTH, HEIGHT, false, objData.content[i].smallTitle || " ", FONT_SIZE, "#000000");
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
   * 创建模板3-3 章节
   */
  createChapters(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[2].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    insertAutoFitText(
      newSlide,
      objData.chapters<10?`0${objData.chapters}`:objData.chapters,
      310, 55, 100, 50, 42, 
      "#0664A5",
      true,
      true
    )
    
    insertAutoFitText(
      newSlide,
      objData.title,
      120, 180, 450, 150, 34, 
      "#ffffff",
      true,
      true
    )

  },

  /**
   * 创建正文-> 1标题1插图(随机左右)
   */
  create1Title1IllustrationLayout(newSlide, presentation, objData) {
    //设置背景图片
    const fileId = TEMPLATE3[6].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    const imageUrl = objData.content.find(it => Boolean(it.imageKey))?.imageKey; //"https://cdn.pixabay.com/photo/2024/12/27/14/58/owl-9294302_640.jpg"
    const __keyword = objData.content?.find(it => it.imageKeyWord)?.imageKeyWord;
    //创建标题
    insertAutoFitText(
      newSlide,
      objData.title,
      80, 20, 250, 60, 22, 
      "#ffffff",
      false,
      false
    )

    //创建图左文右布局
    const createLeftLayout = function () {
      //创建标题
      insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle,
        50, 100, 320, 50, 16, 
        "#000000",
        false,
        false,
        true
      )

      //创建内容
      const _textBox = insertAutoFitText(
        newSlide,
        objData.content[0].smallContent,
        50, 150, 300, 220, 14, 
        "#000000",
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
        const loadingGif = insertPlaceholderGif(520, 130, newSlide, 270, 250);
        const urls = searchGoogleImages(__keyword);

        insertAndCropToFixedSizeFromImages(520, 130, newSlide, urls, 270, 250);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(520, 130, newSlide, imageUrl, 270, 250);
      }
    }
    //创建图右文左布局
    const createRightayout = function () {
      //创建标题
      insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle,
        370, 100, 320, 50, 16, 
        "#000000",
        false,
        false
      )

      //创建内容
      const _textBox = insertAutoFitText(
        newSlide,
        objData.content[0].smallContent,
        370, 150, 300, 220, 14, 
        "#000000",
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
        const loadingGif = insertPlaceholderGif(70, 130, newSlide, 270, 250, true);
        const urls = searchGoogleImages(__keyword);

        insertAndCropToFixedSizeFromImages(70, 130, newSlide, urls, 270, 250, true);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(70, 130, newSlide, imageUrl, 270, 250, true);
      }
    }
    Math.random() > 0.5 ? createRightayout() : createLeftLayout();
  },

  /**
   * 创建模板3-4 两个小标题
   */
  create2SamllTitle(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[3].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    insertAutoFitText(
      newSlide,
      objData.title,
      80, 20, 250, 60, 22, 
      "#ffffff",
      false,
      false
    )

  
    let textBox1 = [], textBox2 = []  
    if(objData.content && objData.content[0]){
      const  text1= insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle||" ",
        90, 140, 130, 30, 14, 
        "#000000",
        false,
        false,
        true
      )
      const  text2= insertAutoFitText(
        newSlide,
        objData.content[0].smallContent||" ",
        90, 180, 380, 60, 12, 
        "#000000",
        false,
        false
      )

      //创建链接
      if (objData.content[0].link) {
        appendTextWithHyperlink(text2, objData.content[0].text, objData.content[0].link)
      }

      if (!objData.content[0]?.imageKey && objData.content[0]?.imageKeyWord) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(660, 210, newSlide, 120, 80);
        const urls = searchGoogleImages(objData.content[0]?.imageKeyWord);
        insertAndCropToFixedSizeFromImages(660, 210, newSlide, urls, 120, 80);
        loadingGif.remove();
      } else if (objData.content[0]?.imageKey) {
        insertAndCropToFixedSize(660, 210, newSlide, objData.content[0]?.imageKey, 120, 80)
      }

      textBox1.push(text1)
      textBox2.push(text2)
    }

    if(objData.content && objData.content[1]){
      const text1 = insertAutoFitText(
        newSlide,
        objData.content[1].smallTitle||" ",
        90, 250, 130, 30, 14, 
        "#000000",
        false,
        false,
        true
      )
      const text2 = insertAutoFitText(
        newSlide,
        objData.content[1].smallContent||" ",
        90, 290, 380, 60, 12, 
        "#000000",
        false,
        false
      )

      //创建链接
      if (objData.content[1].link) {
        appendTextWithHyperlink(text2, objData.content[1].text, objData.content[1].link)
      }

      if (!objData.content[1]?.imageKey && objData.content[1]?.imageKeyWord) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(660, 360, newSlide, 120, 80, true);
        const urls = searchGoogleImages(objData.content[1]?.imageKeyWord);
        insertAndCropToFixedSizeFromImages(660, 360, newSlide, urls, 120, 80, true);
        loadingGif.remove();
      } else if (objData.content[1]?.imageKey) {
        insertAndCropToFixedSize(660, 360, newSlide, objData.content[0]?.imageKey, 120, 80, true)
      }

      textBox1.push(text1)
      textBox2.push(text2)
    }
    setMinFontSize(textBox1)
    setMinFontSize(textBox2)
  },

  /**
   * 创建模板3-6   三个小标题
   */
  create3SmallTitle(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[5].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    insertAutoFitText(
      newSlide,
      objData.title,
      80, 20, 250, 60, 22, 
      "#ffffff",
      false,
      false
    )
    
    let textBox1 = [], textBox2 = []  
    if(objData.content && objData.content[0]){
      const text1 = insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle||" ",
        45, 145, 200, 30, 14, 
        "#000000",
        true,
        false,
        true
      )
      const text2 = insertAutoFitText(
        newSlide,
        objData.content[0].smallContent||" ",
        45, 175, 200, 130, 12, 
        "#000000",
        false,
        false,
      )
      
      //创建链接
      if (objData.content[0].link) {
        appendTextWithHyperlink(text2, objData.content[0].text, objData.content[0].link)
      }
      if (!objData.content[0]?.imageKey && objData.content[0]?.imageKeyWord) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(100, 410, newSlide, 140, 80);
        const urls = searchGoogleImages(objData.content[0]?.imageKeyWord);
        insertAndCropToFixedSizeFromImages(100, 410, newSlide, urls, 140, 80);
        loadingGif.remove();
      } else if (objData.content[0]?.imageKey) {
        insertAndCropToFixedSize(100, 410, newSlide, objData.content[0]?.imageKey, 180, 140)
      }
      textBox1.push(text1)
      textBox2.push(text2)
    }

    if(objData.content && objData.content[1]){
      const text1 = insertAutoFitText(
        newSlide,
        objData.content[1].smallTitle||" ",
        260, 85, 200, 30, 14, 
        "#000000",
        true,
        false,
        true
      )
      const text2 = insertAutoFitText(
        newSlide,
        objData.content[1].smallContent||" ",
        260, 115, 200, 130, 12, 
        "#000000",
        false,
        false,
      )
      //创建链接
      if (objData.content[1].link) {
        appendTextWithHyperlink(text2, objData.content[1].text, objData.content[1].link)
      }
      if (!objData.content[1]?.imageKey && objData.content[1]?.imageKeyWord) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(385, 330, newSlide, 140, 80);
        const urls = searchGoogleImages(objData.content[1]?.imageKeyWord);
        insertAndCropToFixedSizeFromImages(385, 330, newSlide, urls, 140, 80);
        loadingGif.remove();
      } else if (objData.content[1]?.imageKey) {
        insertAndCropToFixedSize(385, 330, newSlide, objData.content[1]?.imageKey, 140, 80)
      }
      textBox1.push(text1)
      textBox2.push(text2)
    }

    if(objData.content && objData.content[2]){
      const text1 = insertAutoFitText(
        newSlide,
        objData.content[0].smallTitle||" ",
        480, 145, 200, 30, 12, 
        "#000000",
        true,
        false,
        true
      )
      const text2 = insertAutoFitText(
        newSlide,
        objData.content[0].smallContent||" ",
        475, 175, 200, 130, 12, 
        "#000000",
        false,
        false,
      )
      //创建链接
      if (objData.content[2].link) {
        appendTextWithHyperlink(text2, objData.content[2].text, objData.content[2].link)
      }
      if (!objData.content[2]?.imageKey && objData.content[2]?.imageKeyWord) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(675, 410, newSlide, 140, 80);
        const urls = searchGoogleImages(objData.content[2]?.imageKeyWord);
        insertAndCropToFixedSizeFromImages(675, 410, newSlide, urls, 140, 80);
        loadingGif.remove();
      } else if (objData.content[2]?.imageKey) {
        insertAndCropToFixedSize(675, 410, newSlide, objData.content[2]?.imageKey, 140, 80)
      }
      
      textBox1.push(text1)
      textBox2.push(text2)
    }

    setMinFontSize(textBox1)
    setMinFontSize(textBox2)

  },

  /**
   * 创建模板3-7 4标题4插图
   */
  create4SmallTitle(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[6].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    insertAutoFitText(
      newSlide,
      objData.title,
      80, 20, 250, 60, 22, 
      "#ffffff",
      false,
      false
    )

    let textBox1 = [], textBox2 = []  
    for (let i = 0; i < objData.content.length; i++) {
      const COLOR = "#000000", C_TOP = 100, C_LEFT = 30
      const text1 = insertAutoFitText(
        newSlide,
        objData.content[i].smallTitle||" ",
        C_LEFT+i*170, C_TOP, 150, 40, 16, 
        COLOR,
        true,
        false,
        true
      )
      const text2 = insertAutoFitText(
        newSlide,
        objData.content[i].smallContent||" ",
        C_LEFT+i*170, C_TOP+35, 150, 100, 14, 
        COLOR,
        false,
        false
      )
      
      //创建链接
      if (objData.content[i].link) {
        appendTextWithHyperlink(text2, objData.content[i].text, objData.content[i].link)
      }
      if (!objData.content[i]?.imageKey && objData.content[i]?.imageKeyWord) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(C_LEFT+i*200+(i+1)*28, C_TOP+250, newSlide, 120, 80);
        const urls = searchGoogleImages(objData.content[i]?.imageKeyWord);
        insertAndCropToFixedSizeFromImages(C_LEFT+i*200+(i+1)*28, C_TOP+250, newSlide, urls, 120, 80);
        loadingGif.remove();
      } else if (objData.content[i]?.imageKey) {
        insertAndCropToFixedSize(C_LEFT+i*200+(i+1)*28, C_TOP+250, newSlide, objData.content[i]?.imageKey, 120, 80)
      }
      textBox1.push(text1)
      textBox2.push(text2)
    }
    setMinFontSize(textBox1)
    setMinFontSize(textBox2)
  },

  /**
   * 多个content,带图
   */
  createMoreContent(newSlide, presentation, objData) {
    const fileId = TEMPLATE3[6].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    
    insertAutoFitText(
      newSlide,
      objData.title,
      80, 20, 250, 60, 22, 
      "#ffffff",
      false,
      false
    )

    const WIDTH = objData.content[0]?.imageKey ? 580 : 880, HEIGHT = 340 / objData.content.length, FONT_SIZE = 15, COLOR = "#ffffff", LEFT = 60, TOP = 130, ANSWER_COLOR = "#000000";
    const _textBoxs = [];
    objData.content.map((it, i) => {
      const _textBox = createTextBox(newSlide, LEFT, TOP + i * HEIGHT, WIDTH, HEIGHT, false, it.smallContent, FONT_SIZE, COLOR, false);
      const count = ((it.smallContent || "").match(/\n/g) || []).length;
      _textBoxs.push({
        textBox: _textBox,
        fontSize: FONT_SIZE,
        width: WIDTH,
        rate: objData.language === 'en' ? 0.45 : 0.7,
        height: Math.max(HEIGHT - 30 * count, 44),
        text: objData.content?.[0]?.smallContent,
        language: objData.language
      })

      if (it.link) {
        appendTextWithHyperlink(_textBox, it.text, it.link)
      }
    })
    if (objData.content[0]?.imageKey) {
      insertAndCropToFixedSize(LEFT + WIDTH + 10, TOP, newSlide, objData.content[0]?.imageKey, 260, 260);
    }
    setAutoFit[_textBoxs];
    const _textBox2 = createTextBox(newSlide, LEFT, TOP + 340, WIDTH, 60, false, objData.text || " ", FONT_SIZE, ANSWER_COLOR, false);
    setAutoFit([
      {
        textBox: _textBox2,
        fontSize: FONT_SIZE,
        width: WIDTH,
        rate: objData.language === 'en' ? 0.3 : 0.7,
        height: 60,
        text: objData.text || " ",
        language: objData.language
      }
    ])
  },
  
  /**
    * 創建音頻播放頁面
    */
  createVoicePlay(newSlide, presentation, objData) {
    const fileId = TEMPLATE3[6].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    insertAutoFitText(
      newSlide,
      objData.title,
      80, 20, 250, 60, 22, 
      "#ffffff",
      false,
      false
    )

    //創建標題
    const VOICE_WIDTH = 40, VOICE_HEIGHT = Math.max(400 / objData.content.length, 70), VOICE_FONT_SIZE = 20, ITEM_WIDTH = 750, ITEM_HEIGHT = Math.max(400 / objData.content.length, 70), ITEM_FONT_SIZE = 15, COLOR = "#000000", TOP = 130, LEFT = 60;

    const _textBoxs = [];
    for (let i = 0; i < objData.content.length; i++) {
      const voiceBox = createTextBox(newSlide, LEFT, TOP + i * ITEM_HEIGHT + 5, VOICE_WIDTH, VOICE_HEIGHT, false, "🎵", VOICE_FONT_SIZE, "#000000", true);
      insertVoice(voiceBox, objData.content?.[i]?.ttsId);
      const _textBox = createTextBox(newSlide, LEFT + 60, TOP + i * ITEM_HEIGHT + 5, ITEM_WIDTH, ITEM_HEIGHT, false, objData.content?.[i]?.smallContent, ITEM_FONT_SIZE, COLOR, false);
      _textBoxs.push({
        textBox: _textBox,
        fontSize: ITEM_FONT_SIZE,
        width: ITEM_WIDTH,
        rate: objData.language === 'en' ? 0.3 : 0.7,
        height: ITEM_HEIGHT,
        text: objData.content?.[i]?.smallContent,
        language: objData.language
      })
      if (objData.content?.[i]?.link) {
        appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[i].link)
      }
    }
    setAutoFit(_textBoxs);
  },


  /**
   * 创建模板3-10 结尾
   */
  createEndingLayout(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[9].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    //创建标题
    insertAutoFitText(
      newSlide,
      objData.title,
      50, 100, 300, 150, 34, 
      "#ffffff",
      true,
      true
    )
    
  },
  /**
   * 创建Loading
   */
  createLoadingPage(newSlide, presentation, objData) {
    const fileId = TEMPLATE3[6].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    insertAutoFitText(
      newSlide,
      objData.title,
      80, 20, 250, 60, 22, 
      "#ffffff",
      false,
      false
    )
    const GIF_SIZE = 100, TOP = 100, pageWidth = presentation.getPageWidth();;
    const ofsetLeft = (pageWidth - GIF_SIZE) / 2 + GIF_SIZE + 15, offsetTop = TOP + 60;
    insertPlaceholderGif(ofsetLeft, offsetTop, newSlide, GIF_SIZE, GIF_SIZE);
  },
  

  /**
   * 创建模板3-5
   */
  createLayout3_5(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[4].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    const pageHeight = convertPtToPx(presentation.getPageHeight()), pageWidth = convertPtToPx(presentation.getPageWidth());
    const TOP = 130, LEFT = 50, WIDTH = 400, IMG_WIDTH = pageWidth - WIDTH - LEFT * 2 - 30, FONT_SIZE = 40, HEIGHT = 90, COLOR = "#ffffff";
    // const textBox = createTextBox(newSlide, LEFT, TOP, WIDTH, HEIGHT, false, objData.title, FONT_SIZE, COLOR, true);
    createTextBox(newSlide, 100, 35, 300, 60, false, objData.title, 22, COLOR, false);

    
    if(objData.content && objData.content[0]){
      createTextBox(newSlide, 200, 282, 150, 40, false, objData.content[0].smallTitle||" ", 14, "#000000", true);
      createTextBox(newSlide, 170, 330, 220, 150, false, objData.content[0].smallContent||" ", 14, "#000000", false);
    }

    if(objData.content && objData.content[1]){
      createTextBox(newSlide, 600, 282, 150, 40, false, objData.content[1].smallTitle||" ", 14, "#000000", true);
      createTextBox(newSlide, 570, 330, 220, 150, false, objData.content[1].smallContent||" ", 14, "#000000", false);
    }

  },
  
  /**
   * 创建模板3-8
   */
  createLayout3_8(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[7].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    const pageHeight = convertPtToPx(presentation.getPageHeight()), pageWidth = convertPtToPx(presentation.getPageWidth());
    const TOP = 130, LEFT = 50, WIDTH = 400, IMG_WIDTH = pageWidth - WIDTH - LEFT * 2 - 30, FONT_SIZE = 40, HEIGHT = 90, COLOR = "#ffffff";
    // const textBox = createTextBox(newSlide, LEFT, TOP, WIDTH, HEIGHT, false, objData.title, FONT_SIZE, COLOR, true);
    createTextBox(newSlide, 100, 35, 300, 60, false, objData.title, 22, COLOR, false);

    
    if(objData.content && objData.content[0]){
      createTextBox(newSlide, 105, 218, 125, 40, false, objData.content[0].smallTitle||" ", 12, "#000000", true);
      createTextBox(newSlide, 75, 260, 180, 140, false, objData.content[0].smallContent||" ", 10, "#000000", false);
    }

    if(objData.content && objData.content[1]){
      createTextBox(newSlide, 320, 175, 125, 40, false, objData.content[1].smallTitle||" ", 12, "#000000", true);
      createTextBox(newSlide, 290, 217, 180, 140, false, objData.content[1].smallContent||" ", 10, "#000000", false);
    }

    if(objData.content && objData.content[2]){
      createTextBox(newSlide, 530, 218, 125, 40, false, objData.content[2].smallTitle||" ", 12, "#000000", true);
      createTextBox(newSlide, 500, 260, 180, 140, false, objData.content[2].smallContent||" ", 10, "#000000", false);
    }

    if(objData.content && objData.content[3]){
      createTextBox(newSlide, 745, 175, 125, 40, false, objData.content[3].smallTitle||" ", 12, "#000000", true);
      createTextBox(newSlide, 715, 217, 180, 140, false, objData.content[3].smallContent||" ", 10, "#000000", false);
    }

  },

  /**
   * 创建模板3-9
   */
  createLayout3_9(newSlide, presentation, objData={}){
    //设置背景图片
    const fileId = TEMPLATE3[8].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    setImageToBackground(newSlide, presentation, blob);

    const pageHeight = convertPtToPx(presentation.getPageHeight()), pageWidth = convertPtToPx(presentation.getPageWidth());
    const TOP = 130, LEFT = 50, WIDTH = 400, IMG_WIDTH = pageWidth - WIDTH - LEFT * 2 - 30, FONT_SIZE = 40, HEIGHT = 90, COLOR = "#ffffff";
    // const textBox = createTextBox(newSlide, LEFT, TOP, WIDTH, HEIGHT, false, objData.title, FONT_SIZE, COLOR, true);
    createTextBox(newSlide, 100, 35, 300, 60, false, objData.title, 22, COLOR, false);

    
    if(objData.content && objData.content[0]){
      createTextBox(newSlide, 270, 220, 150, 40, true, objData.content[0].smallTitle||" ", 12, "#000000", true);
      createTextBox(newSlide, 265, 280, 160, 200, false, objData.content[0].smallContent||" ", 10, "#000000", false);

      const imageUrl = objData.content[0].imageKey;
      const __keyword = objData.content[0].imageKeyWord;
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(90, 280, newSlide, 150, 100);
        const urls = searchGoogleImages(__keyword);
        insertAndCropToFixedSizeFromImages(90, 280, newSlide, urls, 150, 100);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(90, 280, newSlide, imageUrl, 150, 100)
      }
    }

    if(objData.content && objData.content[1]){
      createTextBox(newSlide, 675, 220, 150, 40, true, objData.content[1].smallTitle||" ", 12, "#000000", true);
      createTextBox(newSlide, 670, 280, 160, 200, false, objData.content[0].smallContent||" ", 10, "#000000", false);

      const imageUrl = objData.content[1].imageKey;
      const __keyword = objData.content[1].imageKeyWord;
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(485, 280, newSlide, 150, 100);
        const urls = searchGoogleImages(__keyword);
        insertAndCropToFixedSizeFromImages(485, 280, newSlide, urls, 150, 100);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(485, 280, newSlide, imageUrl, 150, 100)
      }
    }
  },
  
}


function _layou3Test() {
  const ID = '1nFOVQ000Gh6VctLdhBXvaGIaQ5zlCJTOkeTTLmzmlMA';
  const presentation = SlidesApp.openById(ID);
  const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  
  const objData = {
    "template": 3,
    "pptId": "1nUbbuC0PY_OUUcOlbzTktQ_972oLS2JLR9BiZUhab18",
    "chapters": 1,
    "language": "en",
    "title": "Alphabet Overview",
    "type": 4,
    "pptType": 1,
    "content": [
        {
            "smallTitle": "Composition of the Alphabet",
            "imageKeyWord": "alphabet letters",
            "smallContent": "The English alphabet is made up of 26 letters, ranging from A to Z. Each letter has a unique sound and usage in word formation."
        },
        {
            "smallTitle": "Vowels and Consonants",
            "imageKeyWord": "vowels and consonants",
            "smallContent": "The alphabet is categorized into vowels and consonants. There are 5 vowels: A, E, I, O, U, and 21 consonants."
        },
        {
            "smallTitle": "Vowels and Consonants",
            "imageKeyWord": "vowels and consonants",
            "smallContent": "The alphabet is categorized into vowels and consonants. There are 5 vowels: A, E, I, O, U, and 21 consonants."
        },
        {
            "smallTitle": "Alphabet Usage",
            "imageKeyWord": "alphabet usage",
            "smallContent": "Each letter can represent different sounds, which can vary depending on their placement in a word or their combination with other letters."
        }
    ],
    "order": 4,
    "__error": "The image you are trying to use is invalid or corrupt."
}
  
      // objData.type： 0，1，2，3
      // layout3.createEndingLayout(newSlide,presentation,objData)
      // layout3.createCoverLayout(newSlide,presentation,objData)
      // layout3.createCatalogLayout(newSlide,presentation,objData)
      // layout3.createChapters(newSlide,presentation,objData)
      // layout3.createVoicePlay(newSlide,presentation,objData)
      
      // layout3.create1Title1IllustrationLayout(newSlide,presentation,objData)
      layout3.create2SamllTitle(newSlide,presentation,objData)
      // layout3.create3SmallTitle(newSlide,presentation,objData)
      // layout3.create4SmallTitle(newSlide,presentation,objData)
      // layout3.createMoreContent(newSlide,presentation,objData)
      
}