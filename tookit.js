const RLog = require("rlog-js");
const rlog = new RLog();
const docx = require("docx");
const fs = require("fs");
const path = require("path");
const xlsx = require("node-xlsx");
const { exit } = require("process");

rlog.info("Start reading config.json...");
// 读取配置文件
const configPath = path.join(__dirname, "config.json");
const configData = fs.readFileSync(configPath, "utf-8");
const config = JSON.parse(configData);
rlog.success("config.json has been read");

// 读取data文件夹下的xls文件
rlog.info("Start reading data/xls files...");
const filesD = fs.readdirSync(path.join(__dirname, "data"));
const xlsxFilesD = filesD.filter((file) => path.extname(file) === ".xls");
let dataD = xlsx.parse(path.join(__dirname, "data", xlsxFilesD[0]));

// 处理dataD中的值日表数据
let dutyTable = {};
// console.log(dataD[0].data);
dataD[0].data.forEach((item) => {
  if (!item[1].includes("G")) {
    return;
  }
  // 筛选出名字
  let names = item.slice(4);
  // 去除名字中的空值
  names = names.filter((name) => name !== "");
  // 写入至dutyTable
  dutyTable[`${item[1]}-${item[2]}`] = names;
});

// 获取当周值日生函数
// 值日生：循环值日生名单，找到当前周的值日生
function getDuty(week) {
  let duty = {};
  for (let key in dutyTable) {
    let names = dutyTable[key];
    let index = (week - 1) % names.length;
    duty[key] = names[index];
  }
  return duty;
}

rlog.info("Start reading xlsx files...");
let week;
// 获取当前文件夹下xls文件名
const files = fs.readdirSync(__dirname);
const xlsxFiles = files.filter((file) => path.extname(file) === ".xls");
// 文件名例:自动化学院第8周校检成绩公示
if (xlsxFilesD.length === 0) {
  rlog.error("No xls files found in the current directory");
  return;
  exit();
}
// 获取文件名中的周数
week = xlsxFiles[0].match(/\d+/g)[0];
rlog.success("xlsx files have been found");

// week转换为中文
function weekToChinese(num) {
  const chineseDigits = [
    "零",
    "一",
    "二",
    "三",
    "四",
    "五",
    "六",
    "七",
    "八",
    "九",
  ];
  const chineseUnits = ["十", "百", "千", "万"];

  if (num === 0) return chineseDigits[0];

  let result = "";
  const digits = String(num).split("").map(Number);

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

  return result.replace(/零+$/, "").replace(/零+/g, "零");
}

dutyTable = getDuty(week);
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

let data = xlsx.parse(path.join(__dirname, xlsxFiles[0]));
rlog.success("Data has been read,start parse files...");

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
          str += `${name}(${dbData[item[1]][name].length})
           `;
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

// 制作中间表
// console.log(data[0].data);
// [
//   [
//     "青岛校区",
//     "GB3",
//     "1单元",
//     "1层",
//     "101",
//     "A",
//     "电气与自动化工程学院",
//     "2024",
//     "自动化",
//     "自动化2024-3",
//     "",
//   ],
// ];
let newData = [];
data[0].data.forEach((item) => {
  if (!item[1].includes("G")) {
    return;
  }
  // 将宿舍编号相同，但班级不同的合并，班级名称以逗号分隔
  let flag = false;
  for (let i = 0; i < newData.length; i++) {
    if (`${newData[i][1]}-${newData[i][4]}` === `${item[1]}-${item[4]}`) {
      newData[i][9] += `,${item[9]}`;
      flag = true;
      break;
    }
  }
  if (!flag) {
    newData.push([item[0], item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9]]);
  }
})

data = [];
// 顺序调整
newData.forEach((item) => {
  // 筛选C/D级宿舍
  if (item[5] === "C" || item[5] === "D") {
    // 班级名缩写
    let className = item[9].replaceAll("电气工程及其自动化", "电气").replaceAll("机器人工程","机").replaceAll("自动化", "自")
    data.push([className,`${item[1]}-${item[4]}`,dutyTable[`${item[1]}-${item[4]}`] || "",item[5],"地面卫生"]);
  }
})



// console.log(data);
for (let i = 0; i < config.length; i++) {
  rlog.log(`Start processing ${config[i].grade} grade...`);
  result = [];
  const grade = config[i].grade;
  data.forEach((item) => {
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
                            text: "D累计次数",
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
