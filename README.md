# node-excel-module
Expose excel functions in a XLSX file as a JavaScript module.

[![npm](https://img.shields.io/npm/v/excel-module.svg)](https://www.npmjs.com/excel-module)
![GitHub top language](https://img.shields.io/github/languages/top/momocow/node-excel-module.svg)
[![semantic-release](https://img.shields.io/badge/%20%20%F0%9F%93%A6%F0%9F%9A%80-semantic--release-e10079.svg)](https://github.com/semantic-release/semantic-release)
[![Gitmoji](https://img.shields.io/badge/gitmoji-%20😜%20😍-FFDD67.svg?style=flat-square)](https://gitmoji.carloscuesta.me/)

> **master**
> ---
> [![Build Status](https://travis-ci.org/momocow/node-excel-module.svg?branch=master)](https://travis-ci.org/momocow/node-excel-module)

> **dev**
> ---
> [![Build Status](https://travis-ci.org/momocow/node-excel-module.svg?branch=dev)](https://travis-ci.org/momocow/node-excel-module)

- [node-excel-module](#node-excel-module)
  - [Introduction](#introduction)
  - [Installation](#installation)
  - [Usage](#usage)
  - [API](#api)
    - [excelModule.from()](#excelmodulefrom)
    - [Workbook](#workbook)
      - [.compile()](#compile)
      - [CellSpec](#cellspec)
      - [APIFactory](#apifactory)
  - [Example](#example)
  - [License](#license)

## Introduction
- XLSX files are read and parsed by [Exceljs](https://github.com/guyonroche/exceljs).
- Formula evaluation is powered by [hot-formula-parser](https://github.com/handsontable/formula-parser).
- Expose specified cell values and functions via an object.
- The exported object is serializable; that is, the exported object can be serialized to strings through libraries like [serialize-javascript](https://github.com/yahoo/serialize-javascript).
- Merged cells and shared formulas are supported.
- The minimum raw data is included into the compiled context. It works like a charm even if a formula requires the result from another formula.
- Cross-sheet reference supported.

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
  - [`APIFactory`](#apifactory) the exported object

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

#### APIFactory
```ts
type APIFactory = () => Record<string, any>
```

When specifying cells, use excel syntax like `A1`, `$B$2`. **Note** that both are all treated as absolute coordinates.


## Example
See [integration tests](./test/integration/excel-module.test.js) for more details.

- **sum.xlsx**
<table style="text-align: center;">
  <tr><th>row\col</th><th>A</th><th>B</th><th>C</th></tr>
  <tr><th>1</th><td>1</td><td>2</td><td>=SUM(A1:B1)</td></tr>
  <tr><th>2</th><td>3</td><td>4</td><td>=SUM(A2:B2)</td></tr>
  <tr><th>3</th><td>5</td><td>6</td><td>=SUM(A3:B3)</td></tr>
  <tr><th>4</th><td colspan="3">=SUM(A1:B1)</td></tr>
</table>

- **index.js**
```js
const SUM_XLSX = 'path/to/sum.xlsx'

async function main () {
  const workbook = await excelModule.from(SUM_XLSX)
  const apiFactory = await workbook.compile({
    data1: {
      type: Number,
      cell: 'A1'
    },
    data2: {
      type: String,
      cell: 'B1'
    },
    sum: {
      type: Function,
      cell: 'C1',
      args: [ 'A1', 'B1' ]
    },
    sumAll: {
      type: Function,
      cell: 'A4',
      args: [
        'A1', 'B1', 'A2', 'B2', 'A3', 'B3'
      ]
    }
  })

  const api = apiFactory()

  assert(api.data1 === 1)
  assert(api.data2 === '2')
  assert(api.sum() === 3)                      // 1 + 2 = 3
  assert(api.sum(3, 4) === 7)                  // 3 + 4 = 7
  assert(api.sumAll() === 21)                  // 1 + 2 + 3 + 4 + 5 + 6 = 21
  assert(api.sumAll(5, 6, 7, 8, 9, 10) === 45) // 5 + 6 + 7 + 8 + 9 + 10 = 45
}

main()
```

Each compiled function contains a context of raw data. The context of the example above is shown as follow.

```json
{
  "1!$A$1": 1,
  "1!$B$1": 2,
  "1!$C$1": "=SUM(1!$A$1:$B$1)",
  "1!$A$4": "=SUM(1!$C$1:$C$3)",
  "1!$C$2": "=SUM(1!$A$2:$B$2)",
  "1!$A$2": 3,
  "1!$B$2": 4,
  "1!$C$3": "=SUM(1!$A$3:$B$3)",
  "1!$A$3": 5,
  "1!$B$3": 6
}
```

## License
[GLWTPL](LICENSE) base on https://github.com/me-shaon/GLWTPL.
