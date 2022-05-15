const xl = require("excel4node");
const path = require("path");
const fs = require("fs");

// const file = './1269.json'
// const sheet = '1269'
// var jsonData = JSON.parse(fs.readFileSync(file))
//
// jsonData.forEach(item => {
//   if (Number(item.照片长度) !== 0 ) {
//     for (let i = 0; i < item.照片长度; i++) {
//       item.imgUrl = item.imgUrl || []
//       item.imgUrl.push(path.join((__dirname, `./images/${sheet}/${item.name}(${i}).jpg`)))
//     }
//
//   } else {
//     item.imgUrl = []
//   }
//
// })

const myStyle = {
  font: {
    size: 16
  },
  alignment: {
    horizontal: 'center',
    vertical: 'center'
  }
};


/* 过滤的条件: 哪些不需要放入 excel 的字段  来自文件的 json */
const EXCLUDE = [
  '照片地址',
  'imgUrl'
]
/* sheetList  需要写入的工作簿 */
var sheetList = ['1123', '1269', '1177', '1201']
/* 产生的文件名称 */
const fileName = `result.xlsx`;

const downloadExcel = async () => {
  // 创建工作簿
  const wb = new xl.Workbook();
  // 加载工作簿 单元格样式
  const style = wb.createStyle(myStyle);
  sheetList.forEach(sheetItem => {
    const file = `./${sheetItem}.json`
    const sheet = sheetItem
    // 读取每一个 json 文件
    var jsonData = JSON.parse(fs.readFileSync(file))
    // 单独处理照片问题, 因为照片地址是一个数组, 所以需要遍历, 并且把照片地址放入 imgUrl, 放入属于自己单独数据中的 imgUrl
    jsonData.forEach(item => {
      if (Number(item.照片长度) !== 0) {
        for (let i = 0; i < item.照片长度; i++) {
          item.imgUrl = item.imgUrl || []
          item.imgUrl.push(path.join((__dirname, `./images/${sheet}/${item.name}(${i}).jpg`)))
        }

      } else {
        item.imgUrl = []
      }

    })

    /**
     * 单独处理每一个工作簿页面
     * @type {Worksheet}
     */
    // 向工作簿添加一个表格, 添加工作簿页
    let ws = wb.addWorksheet(sheetItem);
    // 拿取json 数据中第一个对象, 获取 key 作为工作簿的表头
    var keys = Object.keys(jsonData[0]);
    // 过滤表头不需要插入 excel 的表头项
    var filterKeys = keys.filter(item => {
      return !EXCLUDE.includes(item)
    })

    // 写入表头
    filterKeys.forEach((item, index) => {
      ws.cell(1, index += 1).string(item).style(style);
    })
    // 写入表头对应数据
    jsonData.forEach((item, i) => {
      // ws.cell(3 + 4 * i, 1).number(i + 1).style(style);
      filterKeys.forEach((it, idx) => {
        var idxs = idx + 1
        ws.cell(3 + 4 * i, idxs).string(item[filterKeys[idx]]).style(style);
      })
      /**
       *  如果自己的数据上有 imgUrl 代表有图片那么就要把这个数据写入到自己对应行的 excel 中去,
       *  放入的位置是最后
      */
      if (item.imgUrl.length) {
        item.imgUrl.forEach(imgs => {
          ws.addImage({
            path: imgs,
            type: 'picture',
            position: {
              type: 'twoCellAnchor',
              from: {
                col: keys.length + 2,
                colOff: 0,
                row: 2 + 4 * i,
                rowOff: 0,
              },
              to: {
                col: keys.length + 3,
                colOff: 0,
                row: 9 + 4 * i,
                rowOff: 0,
              },
            },
          });

        })
      }

    });
  })



  // 写入文件
  wb.write(fileName)

};


downloadExcel()




// {
//  1 "序号": 10,
//  2 "收药人是否是本人": "否",
//  3 "(收药人)联系人姓名": "任志芳",
//  4 "收药人额外电话": "15021077613",
//  5 "身份证号码": "321081193910287555",
//  6 "name": "任志芳",
//  7 "address": "1123弄",
//  8 "联系方式": "15021077613",
//  9 "门牌": 57,
// 10  "户室": 303,
//   "照片地址": "[\n  \"https://pubuserqiniu.paperol.cn/158627462_36_q19_1649822516jwejSQ.png?attname=37_19_934287BF-B44B-4C04-A65C-9A0F8E642B7B.png&e=1657616803&token=-kY3jr8KMC7l3KkIN3OcIs8Q4s40OfGgUHr1Rg4D:4quNp0td7WIkoxnH7mAJAVLQ7_g=\"\n]",
// 11  "照片长度": 1,
// 12  "药名": "瑞格列奈片",
// 13  "生产厂家": "中美华东",
// 14  "药品需要数量": 2,
// 15  "药名名称2": "盐酸吡格列酮片",
// 16  "药名名称2生产厂家": "北京福元医药",
// 17  "药名名称2需要数量": 2,
// 18  "药名名称3生产厂家": " ",
// 19  "是否同意调剂": "是",
// 20  "药品运输是否需要冷藏": "否",
// 21  "支付方式": "自费/医保",
// 22  "手头剩余药品够用到哪一天": "2022-04-16",
// 23  "备注": "谢谢年龄80多了还有孝喘出来当医院病危抢救回来。现在又封控楼，儿女又不在身边。万分感谢"
// },
