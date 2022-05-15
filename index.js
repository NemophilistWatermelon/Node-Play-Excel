const xlsx = require('xlsx');
const mkdirp = require('mkdirp');
const request = require('request');
const fs = require('fs');

const workbook = xlsx.readFile('./copy.xlsx');

const sheetNames = workbook.SheetNames; // 返回 ['sheet1', ...]
const sheetNumber = 3
const fileName = './1269.json'
const picFileName = '1269'

const worksheet = workbook.Sheets[sheetNames[sheetNumber]]; // 返回 sheet1


const data = xlsx.utils.sheet_to_json(worksheet);



var Template = function(o) {
  this.序号 = o.序号
  this.收药人是否是本人 = o.收药人是否是本人
  this['(收药人)联系人姓名'] = o['(收药人)联系人姓名']
  this['收药人额外电话'] = o['收药人额外电话']
  this['身份证号码'] = o['身份证号码']
  this.name = o['姓名']
  this.address = o['居住小区']
  this.联系方式 = o['联系方式']
  this.门牌 = o['门牌']
  this.户室 = o['户室']
  this.照片地址 = o['药照片地址']
  this.照片长度 = o['照片长度']
  this.药名 = o['药名']
  this.生产厂家 = o['生产厂家']
  this.药品需要数量 = o['药品需要数量']
  this.药名名称2 = o['药名名称2']
  this['药名名称2生产厂家'] = o['药名名称2生产厂家']
  this['药名名称2需要数量'] = o['药名名称2需要数量']
  this['药名名称3'] = o['药名名称3']
  this['药名名称3生产厂家'] = o['药名名称3生产厂家']
  this['药品名称4'] = o['药名名称4']
  this['药名名称4生产厂家'] = o['药名名称4生产厂家']
  this['药品名称4需要数量'] = o['药名名称4需要数量']
  this['是否同意调剂'] = o['是否同意调剂']
  this['药品运输是否需要冷藏'] = o['药品运输是否需要冷藏']
  this['支付方式'] = o['支付方式']
  this['手头剩余药品够用到哪一天'] = o['手头剩余药品够用到哪一天']
  this['备注'] = o['备注']
}

var caches = []


// 创建文件夹
// mkdirp(dir);

data.forEach((item, index) => {
  var o = {
  }
  o = item
  o.序号 = String(index)
  o.门牌 = String(item.门牌)
  o.户室 = String(item.户室)
  o.药品需要数量 = o.药品需要数量 ? String(item.药品需要数量) : ''
  o.药名名称2需要数量 = o.药名名称2需要数量 ? String(item.药名名称2需要数量) : ''
  o.药名名称3需要数量 = o.药名名称3需要数量 ? String(item.药名名称3需要数量) : ''
  o.药名名称4需要数量 = o.药名名称4需要数量 ? String(item.药名名称4需要数量) : ''

  if (item['图片']) {
    o['药照片地址'] = JSON.stringify(item['图片'].split('，'), null, 2)
    o['照片长度'] = item['图片'].split('，').length

  } else {
    o['药照片地址'] = JSON.stringify([])
    o['照片长度'] = 0
  }
  o.照片长度 = String(item.照片长度)
  o.药名 = item['药品名称'].replace(/\//g, '*')
  o.备注 = item['备注'] || ''
  caches.push(new Template(o))
})


fs.writeFileSync(fileName, JSON.stringify(caches, null, 2), function(err) {
  if (err) {
    return console.log(err);
  }
  console.log("The file was saved!");
});



// 下载图片
var data_1123 = JSON.parse(fs.readFileSync(fileName))
data_1123.map(realData => {
  var 图片地址 = JSON.parse(realData['照片地址'])
  if (图片地址.length) {
    var baseName = realData.name
    图片地址.forEach((imgUrl, index) => {
      try {
        request.head(imgUrl, (err, res, body) => {
          const dir = './images/' + picFileName;
          request(imgUrl).pipe(fs.createWriteStream(dir + "/" + baseName + '('+ index + ')' +'.jpg'));
        })

      } catch (e) {
        console.log('错误', e)
      }

    })
  }
})
