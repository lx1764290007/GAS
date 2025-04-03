function doGet(e = {}) {
  const name = e.parameter?.name || 'Idea Lab' + "-" +new Date().toLocaleString();  // 如果没有提供 'name' 参数，则默认值为 'World'
  const presentation = SlidesApp.create(name);
  const file = DriveApp.getFileById(presentation.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  const firstSlide = presentation.getSlides()[0];
  firstSlide.remove();
  file.moveTo(getSlideFolder());
  return ContentService.createTextOutput(presentation.getId());
}

function doPost(e = JSON.stringify()) {
  // e.parameter 包含 URL 中的参数
  // e.postData.contents 包含 POST 请求的主体内容
  // e.postData.type 包含 POST 数据的 MIME 类型
  const requestData = JSON.parse(e.postData?.contents || e);
  if (requestData?.pptType === 2) {
    //删除所有的ppt
    removeAll(requestData.pptId);
  } else if (requestData?.pptType === 3) {
    //删除最后一个ppt
    removeTheLastOne(requestData.pptId);
  } else if (requestData?.pptType === 4) {
    //更换图片
    onReplaceImage(requestData.presentationId, requestData.position, requestData.url);
  } else {
    return newGenerate(requestData);
  }
  // 构建响应
  const response = {
    "status": 200,
    "message": "Success"
  };
  // 返回 JSON 响应
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 动态创建模板1
 */
function createCoverLayoutTest() {
  const ID = '1w4cKpyrRzwFvf5pYtb_DESWi03Zodmh1zX2oeCYPs3g';
  const presentation = SlidesApp.openById(ID);
  const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  //创建封面
  // createCoverLayout(newSlide,presentation,TEMPLATE[0].url);
  //创建目录
  // createCatalogLayout(newSlide,presentation,TEMPLATE[1].url);
  //创建章节
  //createChapters(newSlide,presentation,TEMPLATE[2].url);
  //创建正文，一图一小标题一内容
  //create1Title1IllustrationLayout(newSlide,presentation,TEMPLATE[3].url)
  //创建正文 1大标题3小标题1图
  //create1Title3SmallTitle1IllustrationLayout(newSlide,presentation,TEMPLATE[3].url)
  //创建正文一大标题三小标题布局
  // create1Title3SmallTitleLayout(newSlide,presentation,TEMPLATE[5].url)
  //创建正文一大标题四小标题布局
  //create1Title4SmallTitleLayout(newSlide,presentation,TEMPLATE[9].url);
  //创建结尾
  // createEndingLayout(newSlide, presentation, TEMPLATE[11].url)
}

const __testData = [
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":0,"language":"en","type":1,"title":"History of Hong Kong","pptType":1,"contentType":0,"picture":"https://upload.wikimedia.org/wikipedia/commons/3/39/Interior_view_-_Hong_Kong_Museum_of_History_-_DSC00986.JPG","content":[],"order":1},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":0,"language":"en","type":2,"title":"Table of Contents","pptType":1,"contentType":0,"content":[{"smallTitle":"1. Vocabulary Bank","imageKeyWord":"","imageKey":"","smallContent":""},{"smallTitle":"2. Reading and Comprehension","imageKeyWord":"","imageKey":"","smallContent":""},{"smallTitle":"3. Grammar/Sentence Structure","imageKeyWord":"","imageKey":"","smallContent":""},{"smallTitle":"4. Writing and Composition","imageKeyWord":"","imageKey":"","smallContent":""},{"smallTitle":"5. Listening and Dictation","imageKeyWord":"","imageKey":"","smallContent":""}],"order":2},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":2,"language":"en","construct":1,"title":"Reading comprehension","type":7,"contentType":0,"pptType":1,"content":[{"smallTitle":"Understanding Historical Contexts","imageKeyWord":"Hong Kong historical events","smallContent":"In this reading comprehension section, examine the historical contexts behind key events in Hong Kong's history. Analyze texts to extract meaningful insights and improve your understanding of how historical narratives shape the present and future of Hong Kong."},{"smallTitle":"Analyzing Cultural Transformations","imageKeyWord":"Hong Kong cultural transformation","smallContent":"Explore the cultural transformations in Hong Kong's history. This part of the reading comprehension exercises will guide you through significant cultural and societal changes that have influenced the city\u2019s identity."}],"order":2},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":3,"language":"en","title":"Grammar/Sentence Structure: Learning to Use English Connectives","type":3,"contentType":0,"pptType":1,"content":[{"smallTitle":"Learning to Use English Connectives","imageKeyWord":"","smallContent":"This section focuses on how to use English connectives effectively in sentences. Connectives are words that join parts of a sentence, or sentences together, enabling the text to flow smoothly."}],"order":11},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":3,"language":"en","title":"Grammar Connectives","type":4,"contentType":2,"pptType":1,"content":[{"smallTitle":"Introduction to Connectives","imageKeyWord":"English grammar connectives","smallContent":"Connectives are words that link phrases and clauses. They are crucial for the coherence of a sentence.\n\n1. 《And》 is used to connect phrases or clauses that are similar or related.\nExample: I like apples 《and》 oranges.\n\n2. 《Or》 is used to present alternatives or choices.\nExample: Would you like tea 《or》 coffee?\n\n3. 《Both》 is used to indicate that two things are involved or included.\nExample: 《Both》 Alice 《and》 Bob will attend the meeting."}],"order":12},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":1,"language":"en","title":"Vocabulary Bank: Hong Kong's History","type":3,"contentType":0,"pptType":1,"content":[{"smallTitle":"Vocabulary Bank Overview","imageKeyWord":"","smallContent":"This section explores the vocabulary related to Hong Kong's rich and diverse history."}],"order":1},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":3,"language":"en","title":"Connectives Practice","type":4,"contentType":3,"pptType":1,"content":[{"smallTitle":"Fill in the Blanks with Appropriate Connectives","imageKeyWord":"connectives practice","smallContent":"1. She can dance __ sing.\n2. Do you prefer pizza __ pasta?\n3. __ John __ Mary are coming to the party.\n4. We can go to the park __ stay at home.\n5. __ the teacher __ the students were excited about the field trip.\n6. You must choose chocolate __ vanilla."}],"order":13},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":1,"language":"en","title":"Places - Vocabulary","type":4,"contentType":1,"pptType":1,"content":[{"smallTitle":"Victoria Harbour","imageKeyWord":"Victoria Harbour","smallContent":"A famous natural harbor separating Hong Kong Island and Kowloon Peninsula, known for its stunning skyline."},{"smallTitle":"Kowloon","imageKeyWord":"Kowloon","smallContent":"An urban area in Hong Kong known for its vibrant culture and bustling street markets."},{"smallTitle":"Lantau Island","imageKeyWord":"Lantau Island","smallContent":"The largest island in Hong Kong, home to the famous Tian Tan Buddha."},{"smallTitle":"Central","imageKeyWord":"Hong Kong Central","smallContent":"The central business district of Hong Kong Island, known for its skyscrapers and financial institutions."}],"order":4},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":5,"language":"en","construct":1,"title":"Role Play","type":5,"contentType":0,"pptType":1,"content":[{"smallTitle":"Role Play Activity","imageKeyWord":"Hong Kong role play","smallContent":"Engage in a role play about Hong Kong's transition in 1997. One person will play a British official, and another will be a Chinese representative. Discuss the challenges and opportunities of the transition."}],"order":3},

{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":4,"language":"en","title":"Writing and Composition","type":3,"contentType":0,"pptType":1,"content":[{"smallTitle":"Writing Exercises Focused on Hong Kong's History","imageKeyWord":"","smallContent":"This section will delve into the historical aspects of Hong Kong, with a focus on writing and composition exercises. Through these exercises, learners will explore the significant events that have shaped Hong Kong's development and identity."}],"order":14},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":4,"language":"en","title":"Writing - Fill in the Blanks","type":4,"contentType":4,"pptType":1,"content":[{"smallTitle":"Fill in the Blanks: Hong Kong's History","imageKeyWord":"Hong Kong handover","text":"Victoria Harbour, Central, traditional, economic prosperity, cultural diversity, systems, administrative, economic, thriving, impressive, dull, harbor, urbanization, bustling","smallContent":"Hong Kong was a British colony, but in 1997, it was handed over to China. This period marked a significant change in its history. During the colonial era, ________ and ________ were developed as key economic centers. The handover was accompanied by ________ ceremonies, symbolizing the transition of power. Today, Hong Kong enjoys both ________ and ________. The political framework established after the handover is known as 'one country, two ________'. The return of Hong Kong also marked a new era for its legal and ________ systems. Despite being part of China, Hong Kong maintains a high degree of ________ autonomy. Understanding these historical changes is crucial for appreciating today's Hong Kong, a city known for its ________ and ________ skyline."}],"order":15},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":1,"language":"en","title":"Events - Vocabulary","type":4,"contentType":1,"pptType":1,"content":[{"smallTitle":"Handover","imageKeyWord":"Hong Kong handover","smallContent":"The transfer of sovereignty over Hong Kong from the United Kingdom to China in 1997."},{"smallTitle":"Colonial","imageKeyWord":"Hong Kong colonial period","smallContent":"Relating to the period of British rule in Hong Kong from 1841 to 1997."},{"smallTitle":"Treaty","imageKeyWord":"Treaty of Nanking","smallContent":"Refers to various treaties such as the Treaty of Nanking which ceded Hong Kong to Britain."},{"smallTitle":"Reunification","imageKeyWord":"Hong Kong reunification","smallContent":"The process of Hong Kong becoming a part of China again in 1997."}],"order":5},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":2,"language":"en","construct":0,"title":"Reading comprehension","type":7,"contentType":0,"pptType":1,"content":[{"imageKeyWord":"","smallTitle":"Reading comprehension","imageKey":"https://dreamlab.oss-us-east-1.aliyuncs.com/default/readingComprehension.png","link":"https://test.idea.lab.garlicfit.com/game/AIRead_WebGL/?subjectId=1813&origin=https%3A%2F%2Ftest.idea.lab.garlicfit.com","smallContent":"Reading comprehension"}],"order":2},
{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":1,"language":"en","title":"Culture - Vocabulary","type":4,"contentType":1,"pptType":1,"content":[{"smallTitle":"Dim Sum","imageKeyWord":"Dim Sum","smallContent":"Traditional Cantonese cuisine comprising small dishes served with tea."},{"smallTitle":"Cantonese","imageKeyWord":"Cantonese language","smallContent":"A Chinese dialect spoken in Hong Kong, influencing its cultural identity."},{"smallTitle":"Dragon Boat","imageKeyWord":"Dragon Boat Festival","smallContent":"A traditional Chinese sport and cultural event held during the Dragon Boat Festival."},{"smallTitle":"Lantern Festival","imageKeyWord":"Lantern Festival","smallContent":"A traditional Chinese festival celebrated on the fifteenth day of the first month in the lunisolar calendar."}],"order":6},

{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":5,"language":"en","construct":1,"title":"iQuiz","type":6,"contentType":0,"pptType":1,"content":[{"smallTitle":"iQuiz Questions","imageKeyWord":"Hong Kong quiz","smallContent":"Participate in a quiz testing your knowledge on Hong Kong's history, legal systems, and cultural influences. Answer questions to assess your understanding."}],"order":4},

{"template":3,"pptId":"1RPBccY4ntI9w1peJFoigPROLl7WA1AmuyDUpS_uVR6M","chapters":1,"language":"en","title":"Economy - Vocabulary","type":4,"contentType":1,"pptType":1,"content":[{"smallTitle":"Finance","imageKeyWord":"Hong Kong finance","smallContent":"A major sector in Hong Kong's economy, characterized by banking and investment services."},{"smallTitle":"Trade","imageKeyWord":"Hong Kong trade","smallContent":"Integral to Hong Kong's economy due to its strategic location and free port status."},{"smallTitle":"Stock Market","imageKeyWord":"Hong Kong stock market","smallContent":"The financial marketplace where securities are bought and sold, with the Hong Kong Stock Exchange being a significant player."},{"smallTitle":"Port","imageKeyWord":"Hong Kong port","smallContent":"An essential component of Hong Kong's economy, serving as a major logistics and shipping hub."}],"order":7}]
/**
 * 创建模板
 */
function newGenerate(data) {

  const ID = data.pptId;
  const presentation = SlidesApp.openById(ID);
  const newSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  try {
    setOrder(newSlide, data.order);
    useSelector(data, presentation)(newSlide, presentation, data);
    if (data.type === 0) {
      setSequence(presentation);
    }
    const response = {
      "status": 200,
      "message": "Success"
    };
    // 返回 JSON 响应
    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    Logger.log(e)
    logToDrive(JSON.stringify({...data, __error: e.message}));
    if (newSlide) {
      // newSlide.remove();
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 500,
          message: e.message
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

}

function myTest2() {
  //__testData.map(it => newGenerate(it));
  const __end = {
        "template": 3,
        "pptId": "1bnBWugF1vY7M_GuRfjVSbpv-CP5qHCoi8Ir_hGK-u5w",
        "chapters": 2,
        "language": "en",
        "construct": 1,
        "title": "Reading comprehension",
        "type": 7,
        "contentType": 0,
        "pptType": 1,
        "content": [
                {
                        "smallTitle": "Reading comprehension",
                        "imageKeyWord": "critical reading skills",
                        "smallContent": "Read the following passage carefully and answer the questions below to test your understanding. Pay attention to the main ideas, supporting details, and any inferential questions that may arise. This exercise aims to enhance your ability to interpret and analyze text critically."
                }
        ],
        "order": 7
}
  doPost(JSON.stringify(__end))
}

