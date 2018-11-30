const test = require('ava')
const fs = require('fs')
const path = require('path')

const excelModule = require('../..')

const SUM_XLSX = path.join(__dirname, '..', 'fixtures', 'sum.xlsx')


test('sum.xlsx', async t => {
  t.plan(2)

  const workbook = await excelModule.from(SUM_XLSX)
  const api = workbook.compile({
    sum: {
      cell: 'C1',
      args: {
        num1: 'A1',
        num2: 'B1'
      }
    }
  })

  t.is(api.sum(), 3)
  t.is(api.sum(3, 4), 7)
})

test('API Functions should be serializable', async t => {
  t.plan(2)

  const workbook = await excelModule.from(SUM_XLSX)
  const api = workbook.compile({
    sum: {
      cell: 'C1',
      args: {
        num1: 'A1',
        num2: 'B1'
      }
    }
  })

  const newSum = new Function (api.sum.toString())
  t.is(newSum(), 3)
  t.is(newSum(3, 4), 7)
})
