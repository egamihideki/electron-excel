
const xlsx = require('xlsx');

let workbook = xlsx.readFile('test.xlsx');

let names = workbook.Workbook.Names;

// [ { Name: 'itemA', Ref: 'Sheet1!$C$2' },
//   { Name: 'listB', Ref: 'Sheet1!$B$5:$B$7' } ]

for (let i=0; i<names.length; i++) {

  let item = names[i].Ref.split('!');

  let worksheet = workbook.Sheets[item[0]];

  let cells = item[1].split(':');

  if (cells.length === 1) {

    // $を抜く
    let cell = cells[0].replace('$', '');
    cell = cell.replace('$', '');

    console.log(names[i].Name + ':' + worksheet[cell].v);

  } else {

    // 列を取得  $B$5 のとき [1]がB、[2]が5
    let cols = cells[0].split('$');
    let col = cols[1];

    // 行の開始終了を取得
    let rawFrom = cols[2];

    cols = cells[1].split('$');
    let rawTo = cols[2];

    let str = names[i].Name + ':[';

    for (let j=rawFrom; j<=rawTo; j++) {
      str += worksheet[col + j].v + ',';
    }
    str = str.slice(0,-1);
    str += ']';
    console.log(str);
  }
}

