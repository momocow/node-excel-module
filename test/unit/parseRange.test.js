const test = require('ava')

const { parseRange } = require('../../dist/lib/utils')

test('parseRange()', t => {
  const start = {
    row: {
      index: 0 // 0-based
    },
    column: {
      label: 'A'
    }
  }

  const end = {
    row: {
      index: 0 // 0-based
    },
    column: {
      label: 'E'
    }
  }

  t.deepEqual(parseRange(start, end), [
    'A1', 'B1', 'C1', 'D1', 'E1'
  ])
})
