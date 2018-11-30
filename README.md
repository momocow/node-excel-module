# node-excel-module
Expose excel functions in a XLSX file as a JavaScript module.


[![npm](https://img.shields.io/npm/v/excel-module.svg)](https://www.npmjs.com/excel-module)
![GitHub top language](https://img.shields.io/github/languages/top/momocow/node-excel-module.svg)
[![semantic-release](https://img.shields.io/badge/%20%20%F0%9F%93%A6%F0%9F%9A%80-semantic--release-e10079.svg)](https://github.com/semantic-release/semantic-release)
[![Gitmoji](https://img.shields.io/badge/gitmoji-%20😜%20😍-FFDD67.svg?style=flat-square)](https://gitmoji.carloscuesta.me/)

- [node-excel-module](#node-excel-module)
  - [Installation](#installation)
  - [Example](#example)

## Installation
```
npm install excel-module
```

## Example

- `a.xlsx`
  | Row\Col | `A` | `B` | `C`         |
  | ------- | --- | --- | ----------- |
  | `1`     | 1   | 2   | =SUM(A1:A2) |

- `sum.js`
```js
const excelModule = require('excel-module')

// 1. Given a xlsx file path to read
const workbook = excelModule.from('./a.xlsx')

// OR 2. Given a Readable stream to read
const fs = require('fs')
const workbook = excelModule.from(fs.createReadStream('./a.xlsx'))

// Provide a spec of your excel API for compilation
// Here in the spec of this example,
// we define a function named `sum` which is the formula declared in C1
// the 1st argument of `sum` is mapped to A1
// and the 2nd argument of `sum` is mapped to A2 respectively
const excelAPI = workbook.compile({
  sum: {
    cell: 'C1',
    args: [ 'A1', 'B1' ]
  }
}, {
  exposeCells: true
})

// Using default args
assert(sum() === 3)

// Using the args mappings
// now A1 is 3 and B1 is 4 ONLY in this calculation (=SUM(A1:A2))
// the underlying data remains unchanged
assert(sum(3, 4) === 7)

// raw data (since `exposeCells === true`)
assert(excelAPI.A1 === 1)
assert(excelAPI.B1 === 2)
assert(excelAPI.C1 === '=SUM(A1:A2)')
```
