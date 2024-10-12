const RLog = require("rlog-js");
const rlog = new RLog();
const docx = require("docx");
const fs = require("fs");
const path = require("path");
const xlsx = require("node-xlsx");
const docxPdf = require('docx-pdf');
const { exit } = require("process");

// docx转pdf函数
function wordToPdf(filePath) {
  const outputPdfPath = path.join(
      path.dirname(filePath), 
      `${path.basename(filePath, path.extname(filePath))}.pdf`
  );

  docxPdf(filePath, outputPdfPath, (err, result) => {
      if (err) {
          rlog.log('Error during conversion:', err);
      } else {
          rlog.log('PDF file created successfully:', `${path.basename(filePath, path.extname(filePath))}.pdf`);
      }
  });
}


rlog.info("Start reading xlsx files...");

const config = [
  {
    grade: 23,
    teacher: "万老师、秦老师",
  },
  {
    grade: 21,
    teacher: "李老师、骆老师",
  },
];
let week = 5;

// week转换为中文
function weekToChinese(num) {
  const chineseDigits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const chineseUnits = ['十', '百', '千', '万'];

  if (num === 0) return chineseDigits[0];

  let result = '';
  const digits = String(num).split('').map(Number);

  for (let i = 0; i < digits.length; i++) {
      const digit = digits[i];
      const unitIndex = digits.length - 2 - i;

      if (digit !== 0) {
          result += chineseDigits[digit];
          if (unitIndex >= 0) result += chineseUnits[unitIndex];
      } else {
          if (!result.endsWith(chineseDigits[0])) result += chineseDigits[0];
      }
  }

  return result.replace(/零+$/, '').replace(/零+/g, '零');
}


week = weekToChinese(week);

rlog.log("week:", week);

function firstP(name) {
  return `尊敬的${name}： \n`;
}
function secondP(week) {
  return `        为加强自动化学院学生内务管理，同时也是为了更好的提高同学们的自律性，特将您所带班级第${week}周校检成绩C、D级宿舍和人员名单统计如下，使您更好地了解学生的生活情况，并及时对学生不足之处进行一些指正。`;
}

function lastP() {
  return "        由于以上宿舍得C、D，并直接影响我院在全校的排名，望老师积极督促，在大家的努力下，提高我院成绩。\n";
}

function end() {
  return "自动化学院自律会\n";
}

function getTime() {
  // 获取当前时间
  let date = new Date();
  let year = date.getFullYear();
  let month = date.getMonth();
  let day = date.getDate();
  return `${year}年${month + 1}月${day}日`;
}
// 读取当前文件目录下的xlsx文档
rlog.log("Start reading xlsx files...");
const files = fs.readdirSync(__dirname);
const xlsxFiles = files.filter((file) => path.extname(file) === ".xlsx");

// 读取xlsx文档
if (xlsxFiles.length === 0) {
  rlog.error("No xlsx files found in the current directory");
  return;
  exit();
}
let data = xlsx.parse(path.join(__dirname, xlsxFiles[0]));

// 筛选年级
let result;

// 计算累计次数
function getNums(result, week) {
  // 查看是否存在db.json文件
  let db = fs.readFileSync(path.join(__dirname, "db.json"), "utf-8");
  let dbData = JSON.parse(db);
  // 若不存在，创建
  if (!dbData) {
    dbData = {};
  }

  result.map((item) => {
    // item: [ '电气21-1', 'GB3-648', '邢树升', 'C', '地面卫生' ]
    if (item[3] === "D") {
      // 存储相关信息
      if (dbData[item[1]]) {
        if (dbData[item[1]][item[2]]) {
          if (dbData[item[1]][item[2]].includes(week)) {
            // rlog.warning(`重复违规记录：${item[1]} ${item[2]} ${week}`);
          } else {
            dbData[item[1]][item[2]].push(week);
          }
        } else {
          dbData[item[1]][item[2]] = [week];
        }
      } else {
        dbData[item[1]] = {};
        dbData[item[1]][item[2]] = [week];
      }

      let str = "";

      // 累计次数
      if (dbData[item[1]]) {
        // 列出所有名字
        let names = Object.keys(dbData[item[1]]);
        // 遍历所有名字,输出格式：名字(违规次数)

        names.forEach((name) => {
          str += `${name}(${dbData[item[1]][name].length}) `;
        });
      }
      item[5] = str;
    } else {
      item[5] = " ";
    }
    return item;
  });
  fs.writeFileSync(path.join(__dirname, "db.json"), JSON.stringify(dbData));
  rlog.success("db.json has been saved");
  return result;
}

for (let i = 0; i < config.length; i++) {
  rlog.log(`Start processing ${config[i].grade} grade...`);
  result = [];
  const grade = config[i].grade;
  data[0].data.forEach((item) => {
    if (item.length > 0) {
      item[0].indexOf(grade) !== -1 && result.push(item);
    }
  });
  result = getNums(result, week);
  // 生成docx文档
  const doc = new docx.Document({
    creator: "韩雨昊",
    sections: [
      {
        properties: {},
        children: [
          new docx.Paragraph({
            alignment: docx.AlignmentType.LEFT,
            children: [
              new docx.TextRun({
                text: firstP(config[i].teacher),
                size: 24,
                break: 1,
              }),
              new docx.TextRun({
                text: secondP(week),
                size: 24,
                break: 1,
              }),
            ],
          }),
          new docx.Table({
            alignment: docx.AlignmentType.CENTER,
            width: {
              size: 100,
              type: docx.WidthType.PERCENTAGE,
            },
            rows: [
              new docx.TableRow({
                alignment: docx.AlignmentType.CENTER,
                width: {
                  type: docx.WidthType.AUTO,
                },
                children: [
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        width: {
                          type: docx.WidthType.AUTO,
                        },
                        children: [
                          new docx.TextRun({
                            text: "班级",
                            size: 28,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [
                          new docx.TextRun({
                            text: "宿舍",
                            size: 28,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [
                          new docx.TextRun({
                            text: "值日生",
                            size: 28,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [
                          new docx.TextRun({
                            text: "等级",
                            size: 28,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [
                          new docx.TextRun({
                            text: "扣分原因",
                            size: 28,
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableCell({
                    children: [
                      new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [
                          new docx.TextRun({
                            text: "累计次数",
                            size: 28,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              ...result.map((item) => {
                return new docx.TableRow({
                  children: item.map((cell) => {
                    return new docx.TableCell({
                      children: [
                        new docx.Paragraph({
                          alignment: docx.AlignmentType.CENTER,
                          children: [
                            new docx.TextRun({
                              text: cell.replaceAll(" ", ""),
                              size: 24,
                              bold: cell.includes("G") ? true : false,
                            }),
                          ],
                        }),
                      ],
                    });
                  }),
                });
              }),
            ],
          }),
          new docx.Paragraph({
            children: [
              new docx.TextRun({
                text: lastP(),
                size: 24,
                break: 2, // 添加换行
              }),
            ],
          }),
          new docx.Paragraph({
            alignment: docx.AlignmentType.RIGHT,
            children: [
              new docx.TextRun({
                text: end(),
                size: 24,
                break: 1, // 添加换行
              }),
              new docx.TextRun({
                text: getTime(),
                size: 24,
                break: 1, // 添加换行
              }),
            ],
          }),
        ],
      },
    ],
  });
  // 保存docx文档
  const buffer = docx.Packer.toBuffer(doc);
  buffer.then((buffer) => {
    fs.writeFileSync(
      path.join(__dirname, `${config[i].grade}级第${week}周.docx`),
      buffer
    );
    rlog.success(
      `docx file ${config[i].grade}级第${week}周.docx has been saved`
    );
    // wordToPdf(`${config[i].grade}级第${week}周.docx`);
  });
}
