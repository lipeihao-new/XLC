const fs = require('fs'); // Node.js 文件系统模块
const path = require('path'); // 处理文件路径
const xlsx = require('xlsx'); // Excel 解析库

// 自定义解析函数：处理特定格式的表
function processSheetData(rawData) {
    const result = {
        vehicleInfo: {},
        repairDetails: [],
        partsDetails: [],
        notes: [],
    };

    rawData.forEach((row, index) => {
        if (index === 0) return; // 跳过标题行

        // 判断内容类型并分类
        if (row[0] && row[0].includes("车牌号")) {
            // 车辆基本信息
            const key = row[1]?.replace(/：$/, ""); // 去掉 "：" 后缀
            if (key) result.vehicleInfo[key] = row[2];
        } else if (row[1] && row[1] === "故障原因") {
            // 故障描述
            result.repairDetails.push(row[2]);
        } else if (row[1] && row[1].includes("物流名称")) {
            // 配件明细表头（跳过）
        } else if (row[0] && row[0].match(/\d{4}\.\d{2}\.\d{2}/)) {
            // 配件明细
            result.partsDetails.push({
                date: row[0],
                source: row[1],
                partName: row[2],
                quantity: row[3],
                unitPrice: row[4],
                totalPrice: row[5],
            });
        } else if (!row[0] && row[4]) {
            // 备注
            result.notes.push(row[4]);
        }
    });

    return result;
}

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

                // 将表内容转为二维数组（header: 1）
                const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

                // 数据清洗
                const formattedData = processSheetData(rawData);

                console.log(`表 [${sheetName}] 的整理后内容:`, JSON.stringify(formattedData, null, 2));
            });
        }
    });
}

// 定义 Excel 文件所在的文件夹路径
const folderPath = './excel_files'; // 文件夹路径
readExcelFiles(folderPath);
