
var layout_ = {
  /**
   * 创建标题
   */
  createCoverLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[0].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const TOP = 50, LEFT = 20, WIDTH = 430, IMG_WIDTH = pageWidth - WIDTH - LEFT * 2 - 30, FONT_SIZE = 40, HEIGHT = 90, COLOR = "#745AE4";
    //创建标题
    const textBox = createTextBox(newSlide, LEFT, TOP, WIDTH, HEIGHT, true, objData.title, FONT_SIZE, COLOR, true);

    setAutoFit([{
      textBox: textBox,
      fontSize: FONT_SIZE,
      width: WIDTH,
      height: HEIGHT,
      text: objData.title || "Title",
      rate: objData.language === 'en' ? 0.35 : 0.5,
      language: objData.language
    }]);
    if (!objData.picture) {
      const loadingGif = insertPlaceholderGif(LEFT + WIDTH + 20, 75, newSlide, 200, 200);
      const urls = searchGoogleImages(objData.title);
      insertAndCropToFixedSizeFromImages(LEFT + 110, TOP + HEIGHT, newSlide, urls, 200, 200, true);
      loadingGif.remove();
    } else {
      insertAndCropToFixedSize(LEFT + 100, TOP + HEIGHT + 15, newSlide, objData.picture, 200, 200)
    }
  },
  /**
   * 创建目录布局
   */
  createCatalogLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[1].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    createTitleBox(newSlide, presentation, 20, objData.language === 'hk' ? "目錄" : objData.language === "zh" ? "目录" : "Contents", "#E32062", alignCenter = false)
    const _items = [], { length } = objData.content;
    if (length < 1) return;
    const LEFT = 205, TOP = length <= 5 ? 75 : length > 6 && length < 9 ? 110 : length >= 9 ? 85 : 140, HEIGHT = 40, WIDTH = length > 5 ? 300 : 300 * 2, FONT_SIZE = 26;
    const range = length < 5 ? 5 : length === 6 ? 3 : length === 7 ? 4 : length === 8 ? 4 : 5;
    for (let i = 0; i < length; i++) {
      if (i < range) {
        //创建1-5数字文本框
        const _shape = createCircleBox(newSlide, 155, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)) + 60, 44, 44, true, i + 1, 15, "#FFFFFF").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        _shape.getFill().setSolidFill("#E32062");
        // _shape.setShapeType(SlidesApp.ShapeType.OVAL);

        const _textBox = createTextBox(newSlide, LEFT, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)) + 60, WIDTH, HEIGHT, false, objData.content?.[i]?.smallTitle || " ", FONT_SIZE, "#745AE4");
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
        const _shape = createCircleBox(newSlide, 175 + WIDTH + 50, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + ((i - range) * (HEIGHT + 25)) + 60, 44, 44, i < 9, i + 1, i < 9 ? 15 : 7, "#FFFFFF").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        _shape.getFill().setSolidFill("#E32062");
        //  _shape.setShapeType(SlidesApp.ShapeType.OVAL);
        const _textBox = createTextBox(newSlide, 575, TOP + (length > (i - range) ? 0 : ((HEIGHT + 25) * ((i - range) - length) / 2)) + ((i - range) * (HEIGHT + 25)) + 60, WIDTH, HEIGHT, false, objData.content[i].smallTitle || " ", FONT_SIZE, "#745AE4");
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
   * 創建章節佈局
   */
  createChapters(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[2].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const getChapter = () => {
      let __result = "";
      switch (objData.language) {
        case "en":
          __result = "Chapter " + objData.chapters;
          break;
        case "zh":
          __result = "第" + objData.chapters + "章";
          break;
        default:
          __result = "Chapter " + objData.chapters;

      }
      return __result;
    }
    const ORDER_FONT_SIZE = 16, ORDER_BOX_HEIGHT = 44, ORDER_BOX_WIDTH = 150, ORDER_TOP = 183, ORDER_COLOR = "#FFFFFF", LEFT = 60;
    const _orderBox = createTextBox(newSlide, LEFT, ORDER_TOP, ORDER_BOX_WIDTH, ORDER_BOX_HEIGHT, true, getChapter(), ORDER_FONT_SIZE, ORDER_COLOR, true);
    setAutoFit([
      {
        textBox: _orderBox,
        fontSize: ORDER_FONT_SIZE,
        width: ORDER_BOX_WIDTH,
        rate: objData.language === 'en' ? 0.3 : 0.7,
        height: ORDER_BOX_HEIGHT,
        text: getChapter(),
        language: objData.language
      }
    ]);
    const TEXT_BOX_SIZE = 24, TEXT_BOX_WIDTH = 460, TEXT_BOX_HEIGHT = 200, TEXT_COLOR = "#745AE4";
    const _textBox = createTextBox(newSlide, LEFT, ORDER_TOP + 50, TEXT_BOX_WIDTH, TEXT_BOX_HEIGHT, true, objData.title, TEXT_BOX_SIZE, TEXT_COLOR, false);
    setAutoFit([
      {
        textBox: _textBox,
        fontSize: TEXT_BOX_SIZE,
        width: TEXT_BOX_WIDTH,
        rate: objData.language === 'en' ? 0.3 : 0.7,
        height: TEXT_BOX_HEIGHT,
        text: objData.title,
        language: objData.language
      }
    ]);
  },
  create4ItemsWith4Image(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    //創建標題
    createTitleBox2(newSlide, presentation, objData.title, objData.language);
    const TEXT_BOX_WIDTH = 300, FONT_SIZE = 20, TEXT_BOX_HEIGHT = 44, COLOR = "#745AE4", LEFT = 70, TOP = 110, IMAGE_WIDTH = 160, IMAGE_HEIGHT = 100, COLUMN_GAPS = 500, ROW_GAPS = TOP + TEXT_BOX_HEIGHT + IMAGE_HEIGHT / 2 - 60;
    const _items = [];
    //小標題
    for (let i = 0; i < objData.content.length; i++) {
      const urls = searchGoogleImages(objData.content[i].imageKeyWord);//imageKeyWord

      if (i < 2) {
        const textBox1 = createTextBox(newSlide, LEFT + COLUMN_GAPS * i, TOP, TEXT_BOX_WIDTH, TEXT_BOX_HEIGHT, false, objData.content[i]?.smallTitle, FONT_SIZE, COLOR, true);
        //插入圖片
        insertAndCropToFixedSizeFromImages(LEFT - 20 + (TEXT_BOX_WIDTH - IMAGE_WIDTH) / 2 + COLUMN_GAPS * i, TOP + TEXT_BOX_HEIGHT + 10, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT, true);
        _items.push({
          textBox: textBox1,
          fontSize: FONT_SIZE,
          width: TEXT_BOX_WIDTH,
          rate: objData.language === 'en' ? 0.3 : 0.7,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i]?.smallTitle,
          language: objData.language
        });
        if (objData.content?.[i]?.link) {
          appendTextWithHyperlink(textBox1, objData.content[i].text, objData.content[i].link)
        }
      } else {
        const textBox1 = createTextBox(newSlide, LEFT + (COLUMN_GAPS * (i - 2)), TOP + ROW_GAPS + 60, TEXT_BOX_WIDTH, TEXT_BOX_HEIGHT, false, objData.content[i]?.smallTitle, FONT_SIZE, COLOR, true);
        insertAndCropToFixedSizeFromImages(LEFT - 20 + (TEXT_BOX_WIDTH - IMAGE_WIDTH) / 2 + COLUMN_GAPS * (i - 2), TOP + TEXT_BOX_HEIGHT + ROW_GAPS + 70, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT, true);
        _items.push({
          textBox: textBox1,
          fontSize: FONT_SIZE,
          width: TEXT_BOX_WIDTH,
          rate: objData.language === 'en' ? 0.3 : 0.7,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i]?.smallTitle,
          language: objData.language
        });
        if (objData.content?.[i]?.link) {
          appendTextWithHyperlink(textBox1, objData.content[i].text, objData.content[i].link)
        }


      }
    }
    setAutoFit(_items)
  },
  /**
   * 創建閱讀文本
   */
  create1TitleContent(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[4].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    //創建標題
    createTitleBox2(newSlide, presentation, objData.title, objData.language);
    const WIDTH = objData.content[0]?.imageKey ? 580 : 880, HEIGHT = 400, FONT_SIZE = 20, COLOR = "#222222", LEFT = 40, TOP = 105;
    const _textBox = createTextBox(newSlide, LEFT, TOP, WIDTH, HEIGHT, false, objData.content?.[0]?.smallContent, FONT_SIZE, COLOR, false);
    const count = ((objData.content?.[0]?.smallContent || "").match(/\n/g) || []).length;
    setAutoFit([
      {
        textBox: _textBox,
        fontSize: FONT_SIZE,
        width: WIDTH,
        rate: objData.language === 'en' ? 0.3 : 0.7,
        height: Math.max(HEIGHT - 40 * count, 100),
        text: objData.content?.[0]?.smallContent,
        language: objData.language
      }
    ])
    if (objData.content?.[0]?.link) {
      appendTextWithHyperlink(_textBox, objData.content[0].text, objData.content[0].link)
    }
    if (objData.content[0]?.imageKey) {
      insertAndCropToFixedSize(LEFT + WIDTH + 10, TOP, newSlide, objData.content[0]?.imageKey, 300, 300);
    }
  },
  /**
   * 創建填空題加答案
   */
  createBank(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[4].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    //創建標題
    createTitleBox2(newSlide, presentation, objData.title, objData.language);

    const WIDTH = objData.content[0]?.imageKey ? 580 : 880, HEIGHT = 300, FONT_SIZE = 20, COLOR = "#222222", LEFT = 40, TOP = 105, ANSWER_HEIGHT = 100, ANSWER_COLOR = "#234794";
    const _textBox = createTextBox(newSlide, LEFT, TOP, WIDTH, HEIGHT, false, objData.content?.[0]?.smallContent, FONT_SIZE, COLOR, false);
    const count = ((objData.content?.[0]?.smallContent || "").match(/\n/g) || []).length;

    setAutoFit([
      {
        textBox: _textBox,
        fontSize: FONT_SIZE,
        width: WIDTH,
        rate: objData.language === 'en' ? 0.3 : 0.7,
        height: Math.max(HEIGHT - 40 * count, 100),
        text: objData.content?.[0]?.smallContent,
        language: objData.language
      }
    ])
    if (objData.content?.[0]?.link) {
      appendTextWithHyperlink(_textBox, objData.content[0].text, objData.content[0].link)
    }
    if (objData.content[0]?.imageKey) {
      insertAndCropToFixedSize(LEFT + WIDTH + 10, TOP, newSlide, objData.content[0]?.imageKey, 300, 300);
    }
    const _textBox2 = createTextBox(newSlide, LEFT, TOP + HEIGHT + 5, WIDTH, ANSWER_HEIGHT, false, objData.content?.[0]?.text || objData.text || " ", FONT_SIZE, ANSWER_COLOR, false);
    setAutoFit([
      {
        textBox: _textBox2,
        fontSize: FONT_SIZE,
        width: WIDTH,
        rate: objData.language === 'en' ? 0.3 : 0.7,
        height: ANSWER_HEIGHT,
        text: objData.content?.[0]?.text || objData.text || " ",
        language: objData.language
      }
    ])
  },
  /**
   * 創建音頻播放頁面
   */
  createVoicePlay(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[4].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);

    //創建標題
    createTitleBox2(newSlide, presentation, objData.title, objData.language);
    const VOICE_WIDTH = 40, VOICE_HEIGHT = 70, VOICE_FONT_SIZE = 20, ITEM_WIDTH = 750, ITEM_HEIGHT = 70, ITEM_FONT_SIZE = 14, COLOR = "#222222", TOP = 100, LEFT = 60;

    const _textBoxs = [];
    for (let i = 0; i < objData.content.length; i++) {
      const voiceBox = createTextBox(newSlide, LEFT, TOP + i * ITEM_HEIGHT + 5, VOICE_WIDTH, VOICE_HEIGHT, false, "🎵", VOICE_FONT_SIZE, "#234794", true);
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
  createLoadingPage(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[4].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const SMALL_TITLE_SIZE = 30, SMALL_TITLE_WIDTH = 600, TITLE_FONT_COLOR = "#745AE4", TITLE_BOX_HEIGHT = 46, TOP = 100;
    const GIF_SIZE = 100;

    //创建标题
    const titleBox = createTextBox(newSlide, (pageWidth - SMALL_TITLE_WIDTH) / 2 + 120, TOP - 50, SMALL_TITLE_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR, true);
    setAutoFit([{
      textBox: titleBox,
      fontSize: SMALL_TITLE_SIZE,
      width: SMALL_TITLE_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.45 : undefined,
      language: objData.language
    }])
    const ofsetLeft = (pageWidth - GIF_SIZE) / 2 + GIF_SIZE + 15, offsetTop = TOP + 60;

    insertPlaceholderGif(ofsetLeft, offsetTop, newSlide, GIF_SIZE, GIF_SIZE);
  },
  /**
   * 创建结尾
   */
  createEndingLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE2[9].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    //创建大标题
    const TITLE_SIZE = 30, TITLE_BOX_WIDTH = 450, TITLE_BOX_HEIGHT = 80, TITLE_BOX_COLOR = "#745AE4", TOP = 180;
    //创建标题
    const titleBox = createTextBox(newSlide, 70, TOP, TITLE_BOX_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, TITLE_SIZE, TITLE_BOX_COLOR);
    setAutoFit([{
      textBox: titleBox,
      fontSize: TITLE_SIZE,
      width: TITLE_BOX_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.5 : undefined,
      language: objData.language
    }])

  }
}
 