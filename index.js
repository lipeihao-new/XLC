const fs = require('fs'); // Node.js 文件系统模块
const path = require('path'); // 处理文件路径
const xlsx = require('xlsx'); // Excel 解析库

// 定义读取文件夹中所有 Excel 文件的函数
function readExcelFiles(folderPath) {
    // 读取文件夹中的文件
    const files = fs.readdirSync(folderPath);

    files.forEach((file) => {
        // 检查文件是否是 Excel 文件（通过扩展名）
        if (path.extname(file) === '.xlsx' || path.extname(file) === '.xls') {
            console.log(`正在读取文件: ${file}`);

            // 加载 Excel 文件
            const workbook = xlsx.readFile(path.join(folderPath, file));

            // 获取工作表的名字
            const sheetNames = workbook.SheetNames;

            sheetNames.forEach((sheetName) => {
                // 获取表内容
                const sheet = workbook.Sheets[sheetName];

                // 将表内容转为 JSON 数据
                const data = xlsx.utils.sheet_to_json(sheet);

                console.log(`表 [${sheetName}] 的内容:`, data);
            });
        }
    });
}

// 定义 Excel 文件所在的文件夹路径
const folderPath = './excel_files'; // 文件夹路径
readExcelFiles(folderPath);
