# node-excel-module
Expose excel functions in a XLSX file as a JavaScript module.

[![npm](https://img.shields.io/npm/v/excel-module.svg)](https://www.npmjs.com/excel-module)
![GitHub top language](https://img.shields.io/github/languages/top/momocow/node-excel-module.svg)
[![semantic-release](https://img.shields.io/badge/%20%20%F0%9F%93%A6%F0%9F%9A%80-semantic--release-e10079.svg)](https://github.com/semantic-release/semantic-release)
[![Gitmoji](https://img.shields.io/badge/gitmoji-%20😜%20😍-FFDD67.svg?style=flat-square)](https://gitmoji.carloscuesta.me/)

- [node-excel-module](#node-excel-module)
  - [Introduction](#introduction)
  - [Installation](#installation)
  - [Usage](#usage)
  - [API](#api)
    - [excelModule.from()](#excelmodulefrom)
    - [Workbook](#workbook)
      - [.compile()](#compile)
      - [CellSpec](#cellspec)
  - [Example](#example)
    - [More examples](#more-examples)

## Introduction
- XLSX files are read and parsed by [Exceljs](https://github.com/guyonroche/exceljs).
- Formula evaluation is powered by [hot-formula-parser](https://github.com/handsontable/formula-parser).
- Expose specified cell values and functions via an object.
- The exported object is serializable; that is, the exported object can be serialized to strings through libraries like [serialize-javascript](https://github.com/yahoo/serialize-javascript).
- Merged cells and shared formulas are supported.
- The minimum raw data is included into the compiled context. It works like a charm even if a formula requires the result from another formula.

## Installation
```
npm install excel-module
```

## Usage
```js
const excelModule = require('excel-module')
```

## API
### excelModule.from()
- Parameters
  - src `string` | `Readable`
- Return
  - [`Promise<Workbook>`](#workbook)

### Workbook
Workbook is a descendant class inherited from [Exceljs.Workbook](https://github.com/guyonroche/exceljs#create-a-workbook).

#### .compile()
- Parameters
  - spec `Record<string, CellSpec>` The keys of exported object will be the same as the `spec` object, which are names of exported APIs.
- Return
  - `Record<string, any>` the exported object

#### CellSpec
The `type` should be one of the following constructors, `Number`, `Boolean`, `String` and `Function`.

```ts
interface CellSpec {
  cell: string
  type: CellType
  args?: string[]
}

type CellType = |
  FunctionConstructor |
  NumberConstructor |
  StringConstructor |
  BooleanConstructor
```

When specifying cells, use excel syntax like `A1`, `$B$2`. **Note** that both are all treated as absolute coordinates.

> **Caveat**
> -----------------------
> Cross-sheet reference is not yet supported since hot-formula-parser does not support it yet.
> 
> See https://github.com/handsontable/formula-parser/issues/30.


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
    type: Function,
    cell: 'C1',
    args: [ 'A1', 'B1' ]
  }
})

// Using default args
assert(excelAPI.sum() === 3)

// Using the args mappings
// now A1 is 3 and B1 is 4 ONLY in this calculation (=SUM(A1:A2))
// the underlying data remains unchanged
assert(excelAPI.sum(3, 4) === 7)

// raw data (since `exposeCells === true`)
assert(excelAPI.A1 === 1)
assert(excelAPI.B1 === 2)
assert(excelAPI.C1 === '=SUM(A1:A2)')
```
### More examples
You can see [integration tests](https://github.com/momocow/node-excel-module/blob/master/test/integration/excel-module.test.js#L10) for more examples.