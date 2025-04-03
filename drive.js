/**
 * 将Log日志保存至drive中
 */
function logToDrive(logMessage) {
  const logFileName = `${new Date().toLocaleString()}.txt`;
  //找到指定目录
  const folder = DriveApp.getFolderById(LOG_PATH_ID);

  const file = folder.createFile(logFileName, logMessage, MimeType.PLAIN_TEXT);

  Logger.log("日志已写入：" + file.getUrl());
}

/**
 * 根据日期创建slide目录
 */
function getSlideFolder(){
  const name = "slides-" + new Date().toLocaleDateString("zh-CN");
  //按名称查找指定目录
  const folders = DriveApp.getFoldersByName(name);
   // 如果找到，返回第一个文件夹
  if (folders.hasNext()) {
    return folders.next();  // 返回迭代器中的下一个文件夹
  }
  //在指定目录下创建新目录
  const parentFolder = DriveApp.getFolderById(SLIDES_PATH_ID);

  return parentFolder.createFolder(name);
}
