function createSlideTemplate(presentation, config) {
    try {
        // 创建一个新的空白幻灯片
        const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

        // 设置幻灯片背景图
        if (config.backgroundImageUrl) {
            try {
                slide.setBackgroundImage(SlidesApp.FetchImage(config.backgroundImageUrl));
            } catch (e) {
                console.error(`加载背景图片 ${config.backgroundImageUrl} 失败: ${e.message}`);
            }
        }

        // 添加大标题
        addTitle(slide, presentation, config.title);

        // 幻灯片尺寸
        const slideWidth = presentation.getPageWidth();
        const slideHeight = presentation.getPageHeight();

        // 模块的位置和尺寸
        const moduleWidth = (slideWidth - 200) / 2; // 左右各留 100 的边距
        const moduleHeight = (slideHeight - 200) / 2; // 上下各留 100 的边距
        const moduleXOffsets = [100, slideWidth - moduleWidth - 100];
        const moduleYOffset = 150;

        // 遍历每个模块的数据
        const moduleData = config?.modules || [];
        for (let i = 0; i < moduleData.length; i++) {
            const module = moduleData[i];
            const x = moduleXOffsets[i % 2];
            const y = moduleYOffset + Math.floor(i / 2) * (moduleHeight + 50);

            // 验证模块数据格式
            if (!module || typeof module.iconUrl!== 'string' || typeof module.subtitle!== 'string' || typeof module.content!== 'string') {
                throw new Error('模块数据格式不正确。每个模块应包含 iconUrl、subtitle 和 content 属性，且均为字符串类型。');
            }

            // 添加图片 icon
            try {
                const icon = slide.insertImage(module.iconUrl, x, y, 50, 50);
            } catch (e) {
                console.error(`加载图片 ${module.iconUrl} 失败: ${e.message}`);
            }

            // 添加小标题
            addSubtitle(slide, x + 70, y, module.subtitle);

            // 添加文字内容
            addContent(slide, x + 70, y + 40, module.content, moduleWidth - 70, moduleHeight - 40);
        }
    } catch (e) {
        console.error(`创建幻灯片模板时发生错误: ${e.message}`);
    }
}

function addTitle(slide, presentation, titleText) {
    const slideWidth = presentation.getPageWidth();
    const titleWidth = 600;
    let x = (slideWidth - titleWidth) / 2;
    // 确保参数为数字类型
    x = Number(x);
    const title = slide.insertTextBox(titleText, x, 50, titleWidth, 100);
    const titleParagraph = title.getText();
    titleParagraph.getTextStyle().setFontSize(36).setBold(true);
    titleParagraph.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function addSubtitle(slide, x, y, subtitleText) {
    // 确保参数为数字类型
    x = Number(x);
    y = Number(y);
    const subtitle = slide.insertTextBox(subtitleText, x, y, 180, 30);
    subtitle.getText().getTextStyle().setFontSize(24).setBold(true);
}

function addContent(slide, x, y, contentText, width, height) {
    // 确保参数为数字类型
    x = Number(x);
    y = Number(y);
    width = Number(width);
    height = Number(height);
    var content = slide.insertTextBox(contentText, x, y, width, height);
    var contentParagraph = content.getText();
    contentParagraph.getTextStyle().setFontSize(16);

    // 更精确地计算行数
    var fontSize = 16;
    var avgCharWidth = fontSize * 0.6; // 假设平均字符宽度是字体大小的 0.6 倍
    var maxCharsPerLine = Math.floor(width / avgCharWidth);
    var lines = contentText.split('\n');
    var totalLines = 0;
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        totalLines += Math.ceil(line.length / maxCharsPerLine);
    }

    // 根据行数计算新的高度
    var lineHeight = fontSize * 1.2; // 行高为字体大小的 1.2 倍
    var newHeight = totalLines * lineHeight;
    content.setHeight(newHeight);
}

function creatExample() {
    const ID = '1uV_BHGuGHgdFybDZE7tXiBsokJl1tiU5Fdmad9KWYaw';
    const presentation = SlidesApp.openById(ID);

    // 示例调用
    const config = {
        title: "这是大标题",
        backgroundImageUrl: "https://m.media-amazon.com/images/I/71eROBTwkVL.jpg", // 替换为实际的背景图 URL
        modules: [
            {
                iconUrl: "https://dreamlab.oss-us-east-1.aliyuncs.com/ppt/image/image_20250319022348639.png",
                subtitle: "小标题 1",
                content: "这是模块 1 的文字内容。"
            },
            {
                iconUrl: "https://dreamlab.oss-us-east-1.aliyuncs.com/ppt/image/image_20250319022348639.png",
                subtitle: "小标题 2",
                content: "这是模块 2 的文字内容，可能会比较长，这里只是示例。"
            },
            {
                iconUrl: "https://dreamlab.oss-us-east-1.aliyuncs.com/ppt/image/image_20250319022348639.png",
                subtitle: "小标题 3",
                content: "这是模块 3 的文字内容。"
            },
            {
                iconUrl: "https://dreamlab.oss-us-east-1.aliyuncs.com/ppt/image/image_20250319022348639.png",
                subtitle: "小标题 4",
                content: "这是模块 4 的文字内容。"
            },
            {
                iconUrl: "https://dreamlab.oss-us-east-1.aliyuncs.com/ppt/image/image_20250319022348639.png",
                subtitle: "小标题 5",
                content: "这是模块 4 的文字内容。"
            }
        ]
    };

    createSlideTemplate(presentation, config);
}