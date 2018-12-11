const test = require('ava')
const path = require('path')

const excelModule = require('../..')

const FIXTURE_DIR = path.join(__dirname, '..', 'fixtures')
const SUM_XLSX = path.join(FIXTURE_DIR, 'sum.xlsx')
const CROSS_SHEET_XLSX = path.join(FIXTURE_DIR, 'cross-sheet.xlsx')


test('sum.xlsx', async t => {
  t.plan(9)

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

  t.is(api.data1, 1)
  t.is(api.data2, '2')
  t.is(api.sum(), 3)
  t.is(api.sum(3, 4), 7)
  t.is(api.sumAll(), 21)
  t.is(api.sumAll(5, 6, 7, 8, 9, 10), 45)

  // apiFactory function is serializable
  /**
   * @type {typeof api}
   */
  const newAPI = eval(`(${apiFactory.toString()})`)()
  t.not(newAPI.sum, api.sum)
  t.is(newAPI.sum(), 3)
  t.is(newAPI.sum(3, 4), 7)
})

test('cross-sheet.xlsx', async t => {
  t.plan(2)

  const workbook = await excelModule.from(CROSS_SHEET_XLSX)
  const api = (await workbook.compile({
    sum1: {
      type: Function,
      cell: 'B1',
      args: [
        'A1',
        'A2',
        'A3',
        'A4',
        'A5',
        's2!A1',
        's2!A2',
        's2!A3',
        's2!A4',
        's2!A5'
      ]
    },
    sum2: {
      type: Function,
      cell: 's2!B1',
      args: [
        's1!A1',
        's1!A2',
        's1!A3',
        's1!A4',
        's1!A5'
      ]
    }
  }))()

  const nums = []
  for (let i = 0; i < 10; i++) {
    nums.push((i + 1) * 11)
  }
  const total1 = nums.reduce((acc, num) => acc + num, 0)
  const total2 = nums.reduce((acc, num, i) => i < 5 ? acc + num : acc, 0)

  t.is(api.sum1(...nums), total1)
  t.is(api.sum2(...nums), total2)
})
