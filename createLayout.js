var layout1 = {
  /**
 * @desc 创建封面布局
 * @param newSlide ppt页对象，presentation ppt对象，fileId图片fileID
 */
  createCoverLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[0].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const TOP = 170, LEFT = 90, WIDTH = 450, IMG_WIDTH = pageWidth - WIDTH - LEFT * 2 - 30, FONT_SIZE = 40, HEIGHT = 84;
    const textBox = createTextBox(newSlide, LEFT, TOP, WIDTH, HEIGHT, true, objData.title, FONT_SIZE, "#F9E6A2", true);
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
      //先创建一个默认的loading gif
      const loadingGif = insertPlaceholderGif(LEFT + WIDTH + 20, 75, newSlide, 300, 300);
      const urls = searchGoogleImages(objData.title);

      insertAndCropToFixedSizeFromImages(LEFT + WIDTH + 20, 75, newSlide, urls, 300, 300);
      loadingGif.remove();
    } else {
      insertAndCropToFixedSize(LEFT + WIDTH + 20, 75, newSlide, objData.picture, 300, 300)
    }

  },
  /**
 * @desc 创建目录布局
 * @param newSlide ppt页对象，presentation ppt对象，fileId图片fileID
 */

  createCatalogLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[0].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    createTextBox(newSlide, 50, 80, 160, 40, false, objData.language === 'hk' ? "目錄" : objData.language === "zh" ? "目录" : "Contents", 22, "#FFFFFF").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE).getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    const _items = [], { length } = objData.content;
    if (length < 1) return;
    const LEFT = 265, TOP = length <= 5 ? 75 : length > 6 && length < 9 ? 110 : length >= 9 ? 85 : 140, HEIGHT = 40, WIDTH = length > 5 ? 260 : 280 * 2, FONT_SIZE = 26;
    const range = length < 5 ? 5 : length === 6 ? 3 : length === 7 ? 4 : length === 8 ? 4 : 5;
    for (let i = 0; i < length; i++) {
      if (i < range) {
        //创建1-5数字文本框
        createTextBox(newSlide, 205, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)), FONT_SIZE * 2 + 10, HEIGHT, false, i + 1, FONT_SIZE, "#FFFFFF").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        const _textBox = createTextBox(newSlide, LEFT, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + (i * (HEIGHT + 25)), WIDTH, HEIGHT, false, objData.content[i].smallTitle || " ", FONT_SIZE, "#FFFFFF");
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
        createTextBox(newSlide, 520, TOP + (length > range ? 0 : ((HEIGHT + 25) * (range - length) / 2)) + ((i - range) * (HEIGHT + 25)), FONT_SIZE * 2 + 10, HEIGHT, false, i + 1, FONT_SIZE, "#FFFFFF").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
        const _textBox = createTextBox(newSlide, 580, TOP + (length > (i - range) ? 0 : ((HEIGHT + 25) * ((i - range) - length) / 2)) + ((i - range) * (HEIGHT + 25)), WIDTH, HEIGHT, false, objData.content[i].smallTitle || " ", FONT_SIZE, "#FFFFFF");
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
 * 创建一个章节
 */
  createChapters(newSlide, presentation, objData) {
    const fileId = TEMPLATE[2].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const TEXT_BOX_WIDTH = 500, TEXT_BOX_HEIGHT = 60, FONT_SIZE_1 = 40, FONT_SIZE_2 = 24;
    const pageHeight = presentation.getPageHeight(), pageWidth = presentation.getPageWidth();
    //创建章节序号
    const _textBox = createTextBox(newSlide, (pageWidth - TEXT_BOX_WIDTH + 10), (pageHeight - TEXT_BOX_HEIGHT - 80) / 2, TEXT_BOX_WIDTH, TEXT_BOX_HEIGHT, true, objData.chapters, FONT_SIZE_1, "#47A883");
    const textRange1 = _textBox.getText();
    // 设置文本对齐方式为居中
    textRange1.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    //创建章节标题
    const _textBoxTitle = createTextBox(newSlide, (pageWidth - TEXT_BOX_WIDTH + 10), (pageHeight - TEXT_BOX_HEIGHT + 80) / 2, TEXT_BOX_WIDTH, 100, true, objData.title, FONT_SIZE_1, "#47A883").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    setAutoFit([{
      textBox: _textBoxTitle,
      fontSize: FONT_SIZE_1,
      width: TEXT_BOX_WIDTH,
      height: 100,
      text: objData.title || "目录",
      rate: objData.language === 'en' ? 0.6 : undefined,
      language: objData.language
    }]);
    const textRange2 = _textBoxTitle.getText();
    // 设置文本对齐方式为居中
    textRange2.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  },
  /**
   * 创建一个章节
   */
  createChapters(newSlide, presentation, objData) {
    const fileId = TEMPLATE[2].url;
    const file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const TEXT_BOX_WIDTH = 500, TEXT_BOX_HEIGHT = 60, FONT_SIZE_1 = 40, FONT_SIZE_2 = 24;
    const pageHeight = presentation.getPageHeight(), pageWidth = presentation.getPageWidth();
    //创建章节序号
    const _textBox = createTextBox(newSlide, (pageWidth - TEXT_BOX_WIDTH + 10), (pageHeight - TEXT_BOX_HEIGHT - 80) / 2, TEXT_BOX_WIDTH, TEXT_BOX_HEIGHT, true, objData.chapters, FONT_SIZE_1, "#47A883");
    const textRange1 = _textBox.getText();
    // 设置文本对齐方式为居中
    textRange1.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    //创建章节标题
    const _textBoxTitle = createTextBox(newSlide, (pageWidth - TEXT_BOX_WIDTH + 10), (pageHeight - TEXT_BOX_HEIGHT + 80) / 2, TEXT_BOX_WIDTH, 100, true, objData.title, FONT_SIZE_1, "#47A883").setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    setAutoFit([{
      textBox: _textBoxTitle,
      fontSize: FONT_SIZE_1,
      width: TEXT_BOX_WIDTH,
      height: 100,
      text: objData.title || "目录",
      rate: objData.language === 'en' ? 0.6 : undefined,
      language: objData.language
    }]);
    const textRange2 = _textBoxTitle.getText();
    // 设置文本对齐方式为居中
    textRange2.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    // 创建内容
    //通过古诗的学习，培养孩子们的文学素养
    // createTextBox(newSlide, (pageWidth - TEXT_BOX_WIDTH + 10), (pageHeight - TEXT_BOX_HEIGHT + 220) / 2, TEXT_BOX_WIDTH, TEXT_BOX_HEIGHT, false, " ", FONT_SIZE_2, "#000000");

  },
  /**
   * 创建正文-> 1标题1插图(随机左右)
   */
  create1Title1IllustrationLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const imageUrl = objData.content.find(it => Boolean(it.imageKey))?.imageKey; //"https://cdn.pixabay.com/photo/2024/12/27/14/58/owl-9294302_640.jpg"
    const SMALL_TITLE_SIZE = 26, SMALL_TEXT_SIZE = 18, TITLE_FONT_COLOR = "#843C0B", TEXT_FONT_COLOR = "#E69A27", BOX_WIDTH = 430, TITLE_BOX_HEIGHT = 50, TEXT_BOX_HEIGHT = 350, TOP = 80, LEFT = 100, IMAGE_WIDTH = 300, IMAGE_HEIGHT = 300;
    let _items = [], _itemsTitle = [];
    const __keyword = objData.content?.find(it => it.imageKeyWord)?.imageKeyWord;

    //创建图左文右布局
    const createLeftLayout = function () {
      //insertAndCropToFixedSize(LEFT, TOP, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT);
      //创建标题
      const _textBoxTitle = createTextBox(newSlide, LEFT + 30, 40, 700, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR, true);
      _itemsTitle.push({
        textBox: _textBoxTitle,
        fontSize: SMALL_TITLE_SIZE,
        width: BOX_WIDTH + 10,
        height: TITLE_BOX_HEIGHT,
        text: objData.title,
        rate: objData.language === 'en' ? 0.6 : undefined,
        language: objData.language
      })
      //创建内容
      const _textBox = createTextBox(newSlide, LEFT + IMAGE_WIDTH + 50, 40 + TITLE_BOX_HEIGHT + 20, BOX_WIDTH + 10, TEXT_BOX_HEIGHT, false, objData.content[0].smallContent, SMALL_TEXT_SIZE, TEXT_FONT_COLOR);
      _items.push({
        textBox: _textBox,
        fontSize: SMALL_TEXT_SIZE,
        width: BOX_WIDTH + 10,
        height: TEXT_BOX_HEIGHT,
        text: objData.content[0].smallContent,
        rate: objData.language === 'en' ? 0.5 : undefined,
        language: objData.language
      })

      if (objData.content?.[0]?.link) {
        appendTextWithHyperlink(_textBox, objData.content[0].text, objData.content[0].link)
      }
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(LEFT, TOP, newSlide, IMAGE_WIDTH, IMAGE_HEIGHT);
        const urls = searchGoogleImages(__keyword);

        insertAndCropToFixedSizeFromImages(LEFT, TOP+20, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(LEFT, TOP+20, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT);
      }
    }
    //创建图右文左布局
    const createRightayout = function () {

      //创建标题
      const _textBoxTitle = createTextBox(newSlide, LEFT + 30, 40, 700, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR, true);
      _itemsTitle.push({
        textBox: _textBoxTitle,
        fontSize: SMALL_TITLE_SIZE,
        width: BOX_WIDTH + 10,
        height: TITLE_BOX_HEIGHT,
        text: objData.title,
        rate: objData.language === 'en' ? 0.6 : undefined,
        language: objData.language
      })

      //创建内容
      const _textBox = createTextBox(newSlide, LEFT - 30, 40 + TITLE_BOX_HEIGHT + 20, BOX_WIDTH, TEXT_BOX_HEIGHT, false, objData.content[0].smallContent, SMALL_TEXT_SIZE, TEXT_FONT_COLOR);
      //insertPlaceholderGif(LEFT + BOX_WIDTH + 30, TOP, newSlide, IMAGE_WIDTH, IMAGE_HEIGHT);
      //insertAndCropToFixedSize(LEFT + BOX_WIDTH + 30, TOP, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT);
      _items.push({
        textBox: _textBox,
        fontSize: SMALL_TEXT_SIZE,
        width: BOX_WIDTH + 10,
        height: TEXT_BOX_HEIGHT,
        text: objData.content[0].smallContent,
        rate: objData.language === 'en' ? 0.5 : undefined,
        language: objData.language
      })
      if (objData.content?.[0]?.link) {
        appendTextWithHyperlink(_textBox, objData.content[0].text, objData.content[0].link)
      }
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(LEFT + BOX_WIDTH + 30, TOP, newSlide, IMAGE_WIDTH, IMAGE_HEIGHT);
        const urls = searchGoogleImages(__keyword);

        insertAndCropToFixedSizeFromImages(LEFT + BOX_WIDTH + 30, TOP+20, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(LEFT + BOX_WIDTH + 30, TOP+20, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT)
      }
    }

    Math.random() > 0.5 ? createRightayout() : createLeftLayout();
    setAutoFit(_items);
    setAutoFit(_itemsTitle)
  },

  /**
   * 创建正文 1大标题3小标题1图
   */
  create1Title3SmallTitle1IllustrationLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const imageUrl = objData.content.find(it => Boolean(it.imageKey))?.imageKey; //"https://cdn.pixabay.com/photo/2023/11/30/10/52/bears-8421343_640.jpg"; //"https://cdn.pixabay.com/photo/2024/12/27/14/58/owl-9294302_640.jpg"
    const SMALL_TITLE_SIZE = 26, SMALL_TITLE_WIDTH = 700, SMALL_TEXT_SIZE = 18, TITLE_FONT_COLOR = "#843C0B", TEXT_FONT_COLOR = "#E69A27", BOX_WIDTH = 430, TITLE_BOX_HEIGHT = 50, TEXT_BOX_HEIGHT = 40, SMALL_TEXT_BOX_HEIGHT = 90, TOP = 100, LEFT = 90, IMAGE_WIDTH = 300, IMAGE_HEIGHT = 340, TEXT_CONTENT_FONT_SIZE = 18, TEXT_BOX_CONTENT_COLOR = "#767171";
    const __keyword = objData.content?.find(it => it.imageKeyWord)?.imageKeyWord;
    //创建标题
    const titleBox = createTextBox(newSlide, (pageWidth - SMALL_TITLE_WIDTH) / 2 + 120, 40, SMALL_TITLE_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR);
    setAutoFit([{
      textBox: titleBox,
      fontSize: SMALL_TITLE_SIZE,
      width: SMALL_TITLE_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.45 : undefined,
      language: objData.language
    }])
    const textRange = titleBox.getText();

    // 设置文本对齐方式为居中
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    const _items = [], _itemsTitle = [];
    //创建左图右文布局
    const createLeftLayout = function () {

      //插入图片
      //insertAndCropToFixedSize(LEFT, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT);
      //创建内容标题
      for (let i = 0; i < objData.content.length; i++) {
        const _titleBox = createTextBox(newSlide, LEFT + IMAGE_WIDTH + 30, TOP + TITLE_BOX_HEIGHT - 40 + (i * 2 * TEXT_BOX_HEIGHT + i * 50), BOX_WIDTH + 10, TEXT_BOX_HEIGHT, false, objData.content[i].smallTitle, SMALL_TEXT_SIZE, TEXT_FONT_COLOR);
        _itemsTitle.push({
          textBox: _titleBox,
          fontSize: SMALL_TEXT_SIZE,
          width: BOX_WIDTH + 10,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i].smallTitle,
          rate: objData.language === 'en' ? 0.5 : undefined,
          language: objData.language
        })
        //创建内容文字
        const _textBox = createTextBox(newSlide, LEFT + IMAGE_WIDTH + 30, TOP + TITLE_BOX_HEIGHT - 40 + ((i * 2 + 1) * TEXT_BOX_HEIGHT + i * 50), BOX_WIDTH + 10, SMALL_TEXT_BOX_HEIGHT, false, objData.content[i].smallContent, TEXT_CONTENT_FONT_SIZE, TEXT_BOX_CONTENT_COLOR);
        _items.push({
          textBox: _textBox,
          fontSize: TEXT_CONTENT_FONT_SIZE,
          width: BOX_WIDTH + 10,
          height: SMALL_TEXT_BOX_HEIGHT,
          text: objData.content[i].smallContent,
          rate: objData.language === 'en' ? 0.5 : undefined,
          language: objData.language
        })
        if (objData.content?.[i]?.link) {
          appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[i].link)
        }
      }
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(LEFT, TOP + 10, newSlide, IMAGE_WIDTH, IMAGE_HEIGHT);
        const urls = searchGoogleImages(__keyword);
        insertAndCropToFixedSizeFromImages(LEFT, TOP + 10, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(LEFT, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT)
      }
    }
    //创建左文右图布局
    const createRightayout = function () {
      // insertAndCropToFixedSize(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT);
      //创建标题
      //创建内容标题
      for (let i = 0; i < objData.content.length; i++) {
        const _titleBox = createTextBox(newSlide, LEFT - 10, TOP + TITLE_BOX_HEIGHT - 40 + (i * 2 * TEXT_BOX_HEIGHT + i * 50), BOX_WIDTH + 10, TEXT_BOX_HEIGHT, false, objData.content[i].smallTitle, SMALL_TEXT_SIZE, TEXT_FONT_COLOR);
        _itemsTitle.push({
          textBox: _titleBox,
          fontSize: SMALL_TEXT_SIZE,
          width: BOX_WIDTH + 10,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i].smallTitle,
          rate: objData.language === 'en' ? 0.5 : undefined,
          language: objData.language
        })
        //创建内容文字
        const _textBox = createTextBox(newSlide, LEFT - 10, TOP + TITLE_BOX_HEIGHT - 40 + ((i * 2 + 1) * TEXT_BOX_HEIGHT + i * 50), BOX_WIDTH + 10, SMALL_TEXT_BOX_HEIGHT, false, objData.content[i].smallContent, TEXT_CONTENT_FONT_SIZE, TEXT_BOX_CONTENT_COLOR);
        _items.push({
          textBox: _textBox,
          fontSize: TEXT_CONTENT_FONT_SIZE,
          width: BOX_WIDTH + 10,
          height: SMALL_TEXT_BOX_HEIGHT,
          text: objData.content[i].smallContent,
          rate: objData.language === 'en' ? 0.5 : undefined,
          language: objData.language
        })
        if (objData.content?.[i]?.link) {
          appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[i].link)
        }
      }
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, IMAGE_WIDTH, IMAGE_HEIGHT);
        const urls = searchGoogleImages(__keyword);
        insertAndCropToFixedSizeFromImages(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT)
      }
    }

    Math.random() > 0.5 ? createLeftLayout() : createRightayout();
    setAutoFit(_items);
    setAutoFit(_itemsTitle)
  },
  /**
   * 创建正文一大标题二小标题布局
   */
  create1Title2SmallTitleLayout2(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const imageUrl = objData.content.find(it => Boolean(it.imageKey))?.imageKey; //"https://cdn.pixabay.com/photo/2023/11/30/10/52/bears-8421343_640.jpg"; //"https://cdn.pixabay.com/photo/2024/12/27/14/58/owl-9294302_640.jpg"
    const SMALL_TITLE_SIZE = 26, SMALL_TITLE_WIDTH = 700, SMALL_TEXT_SIZE = 24, TITLE_FONT_COLOR = "#843C0B", TEXT_FONT_COLOR = "#E69A27", BOX_WIDTH = 770, TITLE_BOX_HEIGHT = 50, TEXT_BOX_HEIGHT = 40, SMALL_TEXT_BOX_HEIGHT = 140, TOP = 100, LEFT = 90, IMAGE_WIDTH = 300, IMAGE_HEIGHT = 340, TEXT_CONTENT_FONT_SIZE = 18, TEXT_BOX_CONTENT_COLOR = "#767171";

    //创建标题
    const titleBox = createTextBox(newSlide, (pageWidth - SMALL_TITLE_WIDTH) / 2 + 120, 40, SMALL_TITLE_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR);
    setAutoFit([{
      textBox: titleBox,
      fontSize: SMALL_TITLE_SIZE,
      width: SMALL_TITLE_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.5 : undefined,
      language: objData.language
    }])
    const textRange = titleBox.getText();

    // 设置文本对齐方式为居中
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    const _items = [], _itemsTitle = [];
    //创建内容标题
    for (let i = 0; i < objData.content.length; i++) {
      const _titleBox = createTextBox(newSlide, LEFT, TOP + TITLE_BOX_HEIGHT - 40 + (i * 2 * TEXT_BOX_HEIGHT + i * 110), BOX_WIDTH + 10, TEXT_BOX_HEIGHT, false, objData.content[i].smallTitle, SMALL_TEXT_SIZE, TEXT_FONT_COLOR);
      _itemsTitle.push({
        textBox: _titleBox,
        fontSize: SMALL_TEXT_SIZE,
        width: BOX_WIDTH + 10,
        height: TEXT_BOX_HEIGHT,
        text: objData.content[i].smallTitle,
        rate: objData.language === 'en' ? 0.5 : undefined,
        language: objData.language
      })
      //创建内容文字
      const _textBox = createTextBox(newSlide, LEFT, TOP + TITLE_BOX_HEIGHT - 40 + ((i * 2 + 1) * TEXT_BOX_HEIGHT + i * 110), BOX_WIDTH + 10, SMALL_TEXT_BOX_HEIGHT, false, objData.content[i].smallContent, TEXT_CONTENT_FONT_SIZE, TEXT_BOX_CONTENT_COLOR);
      _items.push({
        textBox: _textBox,
        fontSize: TEXT_CONTENT_FONT_SIZE,
        width: BOX_WIDTH + 10,
        height: SMALL_TEXT_BOX_HEIGHT,
        text: objData.content[i].smallContent,
        rate: objData.language === 'en' ? 0.55 : undefined,
        language: objData.language
      })
      if (objData.content?.[i]?.link) {
        appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[i].link)
      }
    }
    setAutoFit(_items);
    setAutoFit(_itemsTitle)
  },
  /**
   * 创建正文一大标题二小标题一图片布局
   */
  create1Title2SmallTitleLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const __keyword = objData.content?.find(it => it.imageKeyWord)?.imageKeyWord;
    const imageUrl = objData.content.find(it => Boolean(it.imageKey))?.imageKey; //"https://cdn.pixabay.com/photo/2023/11/30/10/52/bears-8421343_640.jpg"; //"https://cdn.pixabay.com/photo/2024/12/27/14/58/owl-9294302_640.jpg"
    const SMALL_TITLE_SIZE = 26, SMALL_TITLE_WIDTH = 700, SMALL_TEXT_SIZE = 22, TITLE_FONT_COLOR = "#843C0B", TEXT_FONT_COLOR = "#E69A27", BOX_WIDTH = 430, TITLE_BOX_HEIGHT = 50, TEXT_BOX_HEIGHT = 40, SMALL_TEXT_BOX_HEIGHT = 140, TOP = 100, LEFT = 90, IMAGE_WIDTH = 300, IMAGE_HEIGHT = 340, TEXT_CONTENT_FONT_SIZE = 18, TEXT_BOX_CONTENT_COLOR = "#767171";

    //创建标题
    const titleBox = createTextBox(newSlide, (pageWidth - SMALL_TITLE_WIDTH) / 2 + 120, 40, SMALL_TITLE_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR);
    setAutoFit([{
      textBox: titleBox,
      fontSize: SMALL_TITLE_SIZE,
      width: SMALL_TITLE_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.5 : undefined,
      language: objData.language
    }])
    const textRange = titleBox.getText();

    // 设置文本对齐方式为居中
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    const _items = [], _itemsTitle = [];
    //创建左图右文布局
    const createLeftLayout = function () {

      //插入图片
      //insertAndCropToFixedSize(LEFT, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT);
      //创建内容标题
      for (let i = 0; i < objData.content.length; i++) {
        const _titleBox = createTextBox(newSlide, LEFT + IMAGE_WIDTH + 30, TOP + TITLE_BOX_HEIGHT - 30 + (i * 2 * TEXT_BOX_HEIGHT + i * 110), BOX_WIDTH + 10, TEXT_BOX_HEIGHT, false, objData.content[i].smallTitle, SMALL_TEXT_SIZE, TEXT_FONT_COLOR);
        _itemsTitle.push({
          textBox: _titleBox,
          fontSize: SMALL_TEXT_SIZE,
          width: BOX_WIDTH + 10,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i].smallTitle,
          rate: objData.language === 'en' ? 0.5 : undefined,
          language: objData.language
        })
        //创建内容文字
        const _textBox = createTextBox(newSlide, LEFT + IMAGE_WIDTH + 30, TOP + TITLE_BOX_HEIGHT - 30 + ((i * 2 + 1) * TEXT_BOX_HEIGHT + i * 110), BOX_WIDTH + 10, SMALL_TEXT_BOX_HEIGHT, false, objData.content[i].smallContent, TEXT_CONTENT_FONT_SIZE, TEXT_BOX_CONTENT_COLOR);
        _items.push({
          textBox: _textBox,
          fontSize: TEXT_CONTENT_FONT_SIZE,
          width: BOX_WIDTH + 10,
          height: SMALL_TEXT_BOX_HEIGHT,
          text: objData.content[i].smallContent,
          rate: objData.language === 'en' ? 0.55 : undefined,
          language: objData.language
        })
        if (objData.content?.[i]?.link) {
          appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[i].link)
        }
      }
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(LEFT, TOP + 10, newSlide, IMAGE_WIDTH, IMAGE_HEIGHT);
        const urls = searchGoogleImages(__keyword);
        insertAndCropToFixedSizeFromImages(LEFT, TOP + 10, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(LEFT, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT)
      }
    }
    //创建左文右图布局
    const createRightayout = function () {
      //insertAndCropToFixedSize(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT);
      //创建标题
      //创建内容标题
      for (let i = 0; i < objData.content.length; i++) {
        const _titleBox = createTextBox(newSlide, LEFT, TOP + TITLE_BOX_HEIGHT - 30 + (i * 2 * TEXT_BOX_HEIGHT + i * 110), BOX_WIDTH + 10, TEXT_BOX_HEIGHT, false, objData.content[i].smallTitle, SMALL_TEXT_SIZE, TEXT_FONT_COLOR);
        _itemsTitle.push({
          textBox: _titleBox,
          fontSize: SMALL_TEXT_SIZE,
          width: BOX_WIDTH + 10,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i].smallTitle,
          rate: objData.language === 'en' ? 0.5 : undefined,
          language: objData.language
        })
        //创建内容文字
        const _textBox = createTextBox(newSlide, LEFT, TOP + TITLE_BOX_HEIGHT - 30 + ((i * 2 + 1) * TEXT_BOX_HEIGHT + i * 110), BOX_WIDTH + 10, SMALL_TEXT_BOX_HEIGHT, false, objData.content[i].smallContent, TEXT_CONTENT_FONT_SIZE, TEXT_BOX_CONTENT_COLOR);
        _items.push({
          textBox: _textBox,
          fontSize: TEXT_CONTENT_FONT_SIZE,
          width: BOX_WIDTH + 10,
          height: SMALL_TEXT_BOX_HEIGHT,
          text: objData.content[i].smallContent,
          rate: objData.language === 'en' ? 0.55 : undefined,
          language: objData.language
        })
        if (objData.content?.[i]?.link) {
          appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[i].link)
        }
      }
      if (!imageUrl && __keyword) {
        //先创建一个默认的loading gif
        const loadingGif = insertPlaceholderGif(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, IMAGE_WIDTH, IMAGE_HEIGHT);
        const urls = searchGoogleImages(__keyword);
        insertAndCropToFixedSizeFromImages(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, urls, IMAGE_WIDTH, IMAGE_HEIGHT);
        loadingGif.remove();
      } else if (imageUrl) {
        insertAndCropToFixedSize(LEFT + BOX_WIDTH + 30, TOP + 10, newSlide, imageUrl, IMAGE_WIDTH, IMAGE_HEIGHT)
      }
    }

    Math.random() > 0.1 ? createLeftLayout() : createRightayout();
    setAutoFit(_items);
    setAutoFit(_itemsTitle)
  },

  /**
   * 创建正文一大标题三小标题布局
   */
  create1Title3SmallTitleLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[5].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    //创建大标题
    const TITLE_SIZE = 26, TITLE_BOX_WIDTH = 700, TITLE_BOX_HEIGHT = 50, TITLE_BOX_COLOR = "#843C0B", TOP = 100,
      SMALL_TITLE_SIZE = 20, SMALL_TEXT_BOX_WIDTH = 300, SMALL_TEXT_BOX_HEIGHT = 30, SMALL_TITLE_COLOR = "#FFFFFF";

    //创建标题
    const titleBox = createTextBox(newSlide, (pageWidth - TITLE_BOX_WIDTH) / 2 + 120, 40, TITLE_BOX_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, TITLE_SIZE, TITLE_BOX_COLOR);
    setAutoFit([{
      textBox: titleBox,
      fontSize: TITLE_SIZE,
      width: TITLE_BOX_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.4 : undefined,
      language: objData.language
    }])
    const textRange = titleBox.getText();
    // 设置文本对齐方式为居中
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    const _items = [], _itemsTitle = [], _itemsTitle1 = [];
    for (let i = 0; i < objData.content.length; i++) {
      //创建小标题1
      const smallTitleBox = createTextBox(newSlide, (pageWidth - SMALL_TEXT_BOX_WIDTH) / 2 + 120, TOP + (50 * (i + 1)), SMALL_TEXT_BOX_WIDTH, SMALL_TEXT_BOX_HEIGHT, false, objData.content[i].smallTitle, SMALL_TITLE_SIZE, SMALL_TITLE_COLOR);
      const smallTitleRange = smallTitleBox.getText();
      // 设置文本对齐方式为居中
      smallTitleRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      _itemsTitle1.push({
        textBox: smallTitleBox,
        fontSize: SMALL_TITLE_SIZE,
        width: SMALL_TEXT_BOX_WIDTH,
        height: SMALL_TEXT_BOX_HEIGHT,
        text: objData.content[i].smallTitle,
        rate: objData.language === 'en' ? 0.5 : undefined,
        language: objData.language
      })
    }
    //创建内容
    const CONTENT_FONT_SIZE = 20, CONTENT_TEXT_BOX_WIDTH = 230, CONTENT_TEXT_BOX_HEIGHT = 40, CONTENT_TITLE_COLOR = "#000000", CONTENT_LEFT = 80, CONTENT_TOP = 320;
    for (let i = 0; i < objData.content.length; i++) {
      const contentBox = createTextBox(newSlide, CONTENT_LEFT + 250 * i + 30 * i, CONTENT_TOP, CONTENT_TEXT_BOX_WIDTH, CONTENT_TEXT_BOX_HEIGHT, true, objData.content[i].smallTitle, CONTENT_FONT_SIZE, CONTENT_TITLE_COLOR);
      _itemsTitle.push({
        textBox: contentBox,
        fontSize: CONTENT_FONT_SIZE,
        width: CONTENT_TEXT_BOX_WIDTH,
        height: CONTENT_TEXT_BOX_HEIGHT,
        text: objData.content[i].smallTitle,
        rate: objData.language === 'en' ? 0.6 : undefined,
        language: objData.language
      })
      contentTextRange = contentBox.getText();
      contentTextRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      //创建内容
      const CONTENT_FONT = 20, CONTENT_BOX_HEIGHT = 90, CONTENT_COLOR = "#C59856";
      const _textBox = createTextBox(newSlide, CONTENT_LEFT + 250 * i + 30 * i, CONTENT_TOP + CONTENT_TEXT_BOX_HEIGHT, CONTENT_TEXT_BOX_WIDTH, CONTENT_BOX_HEIGHT, false, objData.content[i].smallContent, CONTENT_FONT, CONTENT_COLOR);
      _items.push({
        textBox: _textBox,
        fontSize: CONTENT_FONT,
        width: CONTENT_TEXT_BOX_WIDTH,
        height: CONTENT_BOX_HEIGHT,
        text: objData.content[i].smallContent,
        rate: objData.language === 'en' ? 0.6 : undefined,
        language: objData.language
      })

      if (objData.content[i]?.link) {
        appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[i].link)
      }
    }
    setAutoFit(_items);
    setAutoFit(_itemsTitle);
    setAutoFit(_itemsTitle1);
  },

  /**
   * 创建正文一大标题三小标题布局
   */
  create1Title3SmallTitleLayout2(newSlide, presentation, objData) {
    const fileId = TEMPLATE[6].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const pageWidth = presentation.getPageWidth();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    //创建大标题
    const TITLE_SIZE = 26, TITLE_BOX_WIDTH = 700, TITLE_BOX_HEIGHT = 50, TITLE_BOX_COLOR = "#843C0B", TOP = 100,
      ORDER_SIZE = 26, ORDER_COLOR = "#FFFFFF",
      SMALL_TITLE_SIZE = 18, SMALL_TEXT_BOX_WIDTH = 520, SMALL_TEXT_BOX_HEIGHT = 28, SMALL_TITLE_COLOR = "#000000",
      SMALL_CONTENT_SIZE = 18, SMALL_CONTENT_BOX_HEIGHT = 28, SMALL_CONTENT_COLOR = "#767171";


    const titleBox = createTextBox(newSlide, (pageWidth - TITLE_BOX_WIDTH) / 2 + 120, 40, TITLE_BOX_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, TITLE_SIZE, TITLE_BOX_COLOR);
    setAutoFit([{
      textBox: titleBox,
      fontSize: TITLE_SIZE,
      width: TITLE_BOX_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.4 : undefined,
      language: objData.language
    }])
    titleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    const _items = [], _itemsTitle = [];
    //创建序号
    for (let i = 0; i < objData.content.length; i++) {
      //创建序号
      createTextBox(newSlide, 160 + i * 50, 330 - 90 * i, 80, 80, true, i + 1, ORDER_SIZE, ORDER_COLOR);
      //创建小标题
      const smallTilteTextBox = createTextBox(newSlide, 280 + i * 50, 350 - 85 * i, SMALL_TEXT_BOX_WIDTH - i * 50, SMALL_TEXT_BOX_HEIGHT, false, objData.content[i].smallTitle, SMALL_TITLE_SIZE, SMALL_TITLE_COLOR);
      smallTilteTextBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      _itemsTitle.push({
        textBox: smallTilteTextBox,
        fontSize: SMALL_TITLE_SIZE,
        width: SMALL_TEXT_BOX_WIDTH - i * 50,
        height: SMALL_TEXT_BOX_HEIGHT,
        text: objData.content[i].smallTitle,
        rate: objData.language === 'en' ? 0.45 : undefined,
        language: objData.language
      })
      //创建内容
      const smallContentTextBox = createTextBox(newSlide, 280 + i * 50, 350 + SMALL_TEXT_BOX_HEIGHT - (85 * i), SMALL_TEXT_BOX_WIDTH - i * 50, SMALL_CONTENT_BOX_HEIGHT, false, objData.content[i].smallContent, SMALL_CONTENT_SIZE, SMALL_CONTENT_COLOR);
      smallContentTextBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      _items.push({
        textBox: smallContentTextBox,
        fontSize: SMALL_CONTENT_SIZE,
        width: SMALL_TEXT_BOX_WIDTH - i * 50,
        height: SMALL_CONTENT_BOX_HEIGHT,
        text: objData.content[i].smallContent,
        rate: objData.language === 'en' ? 0.6 : undefined,
        language: objData.language
      })
      if (objData.content[i]?.link) {
        appendTextWithHyperlink(smallContentTextBox, objData.content[i].text, objData.content[i].link)
      }

    }
    setAutoFit(_items);
    setAutoFit(_itemsTitle);
    // const smallTilteTextBox = createTextBox(newSlide, 280, 350, SMALL_TEXT_BOX_WIDTH, SMALL_TEXT_BOX_HEIGHT, false, "1", SMALL_TITLE_SIZE, SMALL_TITLE_COLOR);
    // smallTilteTextBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
    //   const smallContentTextBox = createTextBox(newSlide, 280, 350 + SMALL_TEXT_BOX_HEIGHT, SMALL_TEXT_BOX_WIDTH, SMALL_CONTENT_BOX_HEIGHT, false, "1", SMALL_CONTENT_SIZE, SMALL_CONTENT_COLOR);
    // smallContentTextBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

    // createTextBox(newSlide, 210, 240, 80, 80, true, "1", ORDER_SIZE, ORDER_COLOR);
    // createTextBox(newSlide, 260, 150, 80, 80, true, "1", ORDER_SIZE, ORDER_COLOR);

  },

  /**
   * 创建正文一大标题四小标题布局3
   */
  create1Title4SmallTitleLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[9].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    //创建大标题
    const TITLE_SIZE = 26, TITLE_BOX_WIDTH = 700, TITLE_BOX_HEIGHT = 50, TITLE_BOX_COLOR = "#843C0B", TOP = 100;
    //创建标题
    const titleBox = createTextBox(newSlide, (pageWidth - TITLE_BOX_WIDTH) / 2 + 120, 40, TITLE_BOX_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, TITLE_SIZE, TITLE_BOX_COLOR, true);

    const textRange = titleBox.getText();
    const _items = [], _itemsTitle = [];
    // 设置文本对齐方式为居中
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    setAutoFit([{
      textBox: titleBox,
      fontSize: TITLE_SIZE,
      width: TITLE_BOX_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.45 : undefined,
      language: objData.language
    }])
    const ORDER_FONT_SIZE = 26, CONTENT_FONT_SIZE = 24, ORDER_BOX_WIDTH = 80, ORDER_BOX_HEIGHT = 80, ORDER_FONT_COLOR = "#FFFFFF", CONTENT_BOX_COLOR = "#000000", ORDER_BOX_TOP = TOP + 75, ORDER_LEFT = 115, CONTENT_BOX_WIDTH = 250, CONTENT_BOX_HEIGHT = 40, TEXT_SIZE = 16, TEXT_COLOR = "#AFABAB", TEXT_BOX_HEIGHT = 100;
    for (let i = 0; i < objData.content.length; i++) {
      if (i < 2) {
        //数字框
        createTextBox(newSlide, ORDER_LEFT, ORDER_BOX_TOP + 154 * i, ORDER_BOX_WIDTH, ORDER_BOX_HEIGHT, true, i + 1, ORDER_FONT_SIZE, ORDER_FONT_COLOR);
        //标题框
        const _textBoxTitle = createTextBox(newSlide, ORDER_LEFT + ORDER_BOX_WIDTH, ORDER_BOX_TOP + 150 * i - 20, CONTENT_BOX_WIDTH, CONTENT_BOX_HEIGHT, false, objData.content[i].smallTitle, CONTENT_FONT_SIZE, CONTENT_BOX_COLOR);
        _itemsTitle.push({
          textBox: _textBoxTitle,
          fontSize: CONTENT_FONT_SIZE,
          width: CONTENT_BOX_WIDTH,
          height: CONTENT_BOX_HEIGHT,
          text: objData.content[i].smallTitle,
          rate: objData.language === 'en' ? 0.55 : undefined,
          language: objData.language
        })
        //内容框
        const _textBox = createTextBox(newSlide, ORDER_LEFT + ORDER_BOX_WIDTH, ORDER_BOX_TOP + 150 * i + 20, CONTENT_BOX_WIDTH, TEXT_BOX_HEIGHT, false, objData.content[i].smallContent, TEXT_SIZE, TEXT_COLOR);
        _items.push({
          textBox: _textBox,
          fontSize: TEXT_SIZE,
          width: CONTENT_BOX_WIDTH,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i].smallContent,
          rate: objData.language === 'en' ? 0.45 : 0.95,
          language: objData.language
        })
        if (objData.content?.[i]?.link) {
          appendTextWithHyperlink(_textBox, objData.content[i].text, objData.content[0].link)
        }
      } else {
        createTextBox(newSlide, ORDER_LEFT + CONTENT_BOX_WIDTH + 150, ORDER_BOX_TOP + 154 * (i - 2), ORDER_BOX_WIDTH, ORDER_BOX_HEIGHT, true, i + 1, ORDER_FONT_SIZE, ORDER_FONT_COLOR);
        //标题框
        const _textBoxTitle = createTextBox(newSlide, ORDER_LEFT + CONTENT_BOX_WIDTH + 150 + ORDER_BOX_WIDTH, ORDER_BOX_TOP + 150 * (i - 2) - 20, CONTENT_BOX_WIDTH, CONTENT_BOX_HEIGHT, false, objData.content[i].smallTitle, CONTENT_FONT_SIZE, CONTENT_BOX_COLOR);
        _itemsTitle.push({
          textBox: _textBoxTitle,
          fontSize: CONTENT_FONT_SIZE,
          width: CONTENT_BOX_WIDTH,
          height: CONTENT_BOX_HEIGHT,
          text: objData.content[i].smallTitle,
          rate: objData.language === 'en' ? 0.55 : undefined,
          language: objData.language
        })
        const _textBox = createTextBox(newSlide, ORDER_LEFT + CONTENT_BOX_WIDTH + ORDER_BOX_WIDTH + 150, ORDER_BOX_TOP + 150 * (i - 2) + 20, CONTENT_BOX_WIDTH, TEXT_BOX_HEIGHT, false, objData.content[i].smallContent, TEXT_SIZE, TEXT_COLOR);
        _items.push({
          textBox: _textBox,
          fontSize: TEXT_SIZE,
          width: CONTENT_BOX_WIDTH,
          height: TEXT_BOX_HEIGHT,
          text: objData.content[i].smallContent,
          rate: objData.language === 'en' ? 0.45 : 0.95,
          language: objData.language
        })
        if (objData.content?.[0]?.link) {
          appendTextWithHyperlink(_textBox, objData.content[0].text, objData.content[0].link);

        }
      }
      setAutoFit(_items);
      setAutoFit(_itemsTitle);
    }

  },
  /**
   * 創建音頻播放頁面
   */
  createVoicePlay(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const SMALL_TITLE_SIZE = 26, SMALL_TITLE_WIDTH = 700, TITLE_FONT_COLOR = "#843C0B", TITLE_BOX_HEIGHT = 50;

    const titleBox = createTextBox(newSlide, (pageWidth - SMALL_TITLE_WIDTH) / 2 + 120, 40, SMALL_TITLE_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR, true);
    setAutoFit([{
      textBox: titleBox,
      fontSize: SMALL_TITLE_SIZE,
      width: SMALL_TITLE_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.45 : undefined,
      language: objData.language
    }])
    const VOICE_WIDTH = 40, VOICE_HEIGHT = Math.max(400 / objData.content.length, 70), VOICE_FONT_SIZE = 20, ITEM_WIDTH = 750, ITEM_HEIGHT = Math.max(400 / objData.content.length, 70), ITEM_FONT_SIZE = 15, COLOR = "#222222", TOP = 100, LEFT = 60;

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
  /**
 * 创建四个图片
 */
  create4ItemsWith4Image(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    //創建標題
    const SMALL_TITLE_SIZE = 26, SMALL_TITLE_WIDTH = 700, TITLE_FONT_COLOR = "#843C0B", TITLE_BOX_HEIGHT = 50;
    const pageWidth = presentation.getPageWidth();

    const titleBox = createTextBox(newSlide, (pageWidth - SMALL_TITLE_WIDTH) / 2 + 120, 40, SMALL_TITLE_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR, true);
    setAutoFit([{
      textBox: titleBox,
      fontSize: SMALL_TITLE_SIZE,
      width: SMALL_TITLE_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.45 : undefined,
      language: objData.language
    }])
    const TEXT_BOX_WIDTH = 300, FONT_SIZE = 20, TEXT_BOX_HEIGHT = 44, COLOR = "#222222", LEFT = 160, TOP = 110, IMAGE_WIDTH = 160, IMAGE_HEIGHT = 100, COLUMN_GAPS = 320, ROW_GAPS = TOP + TEXT_BOX_HEIGHT + IMAGE_HEIGHT / 2 - 60;
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
   * 多个content,带图
   */
  createMoreContent(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    //創建標題
    const SMALL_TITLE_SIZE = 26, SMALL_TITLE_WIDTH = 700, TITLE_FONT_COLOR = "#843C0B", TITLE_BOX_HEIGHT = 50;
    const pageWidth = presentation.getPageWidth();

    const titleBox = createTextBox(newSlide, (pageWidth - SMALL_TITLE_WIDTH) / 2 + 120, 40, SMALL_TITLE_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, SMALL_TITLE_SIZE, TITLE_FONT_COLOR, true);
    setAutoFit([{
      textBox: titleBox,
      fontSize: SMALL_TITLE_SIZE,
      width: SMALL_TITLE_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.45 : undefined,
      language: objData.language
    }])
    const __imageKey = objData.content.find(it => it.imageKeyWord), __imageUrl = objData.content.find(it => it.imageKey);
    const WIDTH = __imageKey || __imageUrl ? 560 : 850, HEIGHT = (objData.text ? 340 : 400) / objData.content.length, FONT_SIZE = Math.max(18 - objData.content.length * 2, 14), COLOR = "#222222", LEFT = 50, TOP = 105, ANSWER_COLOR = "#234794";
    const _textBoxs = [];
    objData.content.map((it, i) => {
      const _textBox = createTextBox(newSlide, LEFT + 10, TOP + i * HEIGHT, WIDTH, HEIGHT, false, it.smallContent, FONT_SIZE, COLOR, false);
      const count = ((it.smallContent || "").match(/\n/g) || []).length;
      _textBoxs.push({
        textBox: _textBox,
        fontSize: FONT_SIZE,
        width: WIDTH,
        rate: objData.language === 'en' ? 0.4 : 0.7,
        height: HEIGHT,
        text: it.smallContent,
        language: objData.language
      })

      if (it.link) {
        appendTextWithHyperlink(_textBox, it.text, it.link)
      }
    })
    if (__imageUrl) {
      insertAndCropToFixedSize(LEFT + WIDTH + 30, 120, newSlide, __imageUrl.imageKey, 260, 260);
    } else if (__imageKey) {
      const urls = searchGoogleImages(__imageKey.imageKey);
      insertAndCropToFixedSizeFromImages(LEFT + WIDTH + 30, 120, newSlide, urls, 260, 260);
    }
    setAutoFit(_textBoxs);
    if (!objData.text) return
    const _textBox2 = createTextBox(newSlide, LEFT, TOP + 340, WIDTH, 60, false, objData.text || " ", FONT_SIZE, ANSWER_COLOR, false);
    setAutoFit([
      {
        textBox: _textBox2,
        fontSize: FONT_SIZE,
        width: WIDTH,
        rate: objData.language === 'en' ? 0.45 : 0.7,
        height: 60,
        text: objData.text || " ",
        language: objData.language
      }
    ])
  },
  /**
   * 创建结尾
   */
  createEndingLayout(newSlide, presentation, objData) {
    const fileId = TEMPLATE[11].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    //创建大标题
    const TITLE_SIZE = 40, TITLE_BOX_WIDTH = 600, TITLE_BOX_HEIGHT = 80, TITLE_BOX_COLOR = "#F9E6A2", TOP = 250;
    //创建标题
    const titleBox = createTextBox(newSlide, (pageWidth - TITLE_BOX_WIDTH) / 2 + 120, TOP - 50, TITLE_BOX_WIDTH, TITLE_BOX_HEIGHT, true, objData.title, TITLE_SIZE, TITLE_BOX_COLOR);
    setAutoFit([{
      textBox: titleBox,
      fontSize: TITLE_SIZE,
      width: TITLE_BOX_WIDTH,
      height: TITLE_BOX_HEIGHT,
      text: objData.title,
      rate: objData.language === 'en' ? 0.5 : undefined,
      language: objData.language
    }])
    const textRange = titleBox.getText();
    // 设置文本对齐方式为居中
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  },
  /**
   * 创建Loading
   */
  createLoadingPage(newSlide, presentation, objData) {
    const fileId = TEMPLATE[3].url;
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    //设置背景图片
    setImageToBackground(newSlide, presentation, blob);
    const pageWidth = presentation.getPageWidth();
    const SMALL_TITLE_SIZE = 24, SMALL_TITLE_WIDTH = 600, TITLE_FONT_COLOR = "#843C0B", TITLE_BOX_HEIGHT = 46, TOP = 100;
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
  }
}
