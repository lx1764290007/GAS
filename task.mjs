import * as fs from 'node:fs'
import * as path from "node:path";

function createFolder(dirPath = "./build") {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath);
    }
}

/**
 * 移动文件
 * @param srcPath
 * @param dstPath
 */
function copyFile(srcPath, dstPath) {
// 读取源目录中的所有文件
    fs.readdirSync(srcPath).forEach(file => {
        const sourceFilePath = path.join(srcPath, file);
        const targetFilePath = path.join(dstPath, file);

        // 检查是否为 .js 文件
        if (fs.statSync(sourceFilePath).isFile() && (path.extname(file) === '.js'|| path.basename(file,".json")==="appsscript")) {
            fs.copyFileSync(sourceFilePath, targetFilePath);

        }
    });
}
function main(){
    createFolder("./build");
    copyFile("./", "./build");
    process.exit(1);
}
main();