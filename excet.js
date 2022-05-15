var excellent = require('excellent');
var fs = require('fs');
var dkGreyBorder = {style: 'thin', color: 'Charcoal Gray'};
var doc = excellent.create({
  sheets: {
    'Summary': {
      rows: [{
        cells: [
          'foo',
          {
            value: 'bar',
          },
          {value: 'foo', style: 'lemonBg'},
          'baz',
          {value: 'quux', style: 'lemonBgBold'},
          {
            value: 'dasdsad',
            image: {image: fs.readFileSync(__dirname + '/images/1123/顾克俭(0).jpg'), filename: '顾克俭(0).jpg'},
          }
        ]
      }]
    }
  },
  styles: {
    borders: [{label: 'dkGrey', left: dkGreyBorder, right: dkGreyBorder, top: dkGreyBorder, bottom: dkGreyBorder}],
    fonts: [{label: 'bold', bold: true}, {label: 'brick', color: 'Brick Red'}],
    fills: [{label: 'lemon', type: 'pattern', color: 'Lemon Glacier'}],
    cellStyles: [
      {label: 'bold', font: 'bold'},
      {label: 'brick', font: 'brick', border: 'dkGrey'},
      {label: 'lemonBg', fill: 'lemon'},
      {label: 'lemonBgBold', font: 'bold', fill: 'lemon'},
      {label: 'dotty', fill: 1}
    ]
  }
});

fs.writeFileSync(__dirname + '/test.xlsx', doc.file);
