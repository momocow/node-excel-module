const test = require('ava')

const { Coordinate } = require('../../dist/lib/coordinate')
const { buildContext, toAbsCoord } = require('../../dist/lib/utils')

test('buildContext()', async t => {
  t.plan(1)

  const table = {
    '1': {
      $A$1: '1',
      $B$1: '2',
      $C$1: '=SUM(D1,E1)',
      $D$1: '10',
      $E$1: '20',
      $F$1: '###',
      $G$1: '***',
      $A$2: '=SUM(A1:E1)'
    }
  }

  const context = await buildContext([ '1!A2' ], (label, done) => {
    const coord = new Coordinate(toAbsCoord(label))
    done(coord.sheet, coord.label, table[coord.sheet][coord.label])
  })

  t.deepEqual(context, {
    1: {
      $A$1: '1',
      $B$1: '2',
      $C$1: '=SUM(D1,E1)',
      $D$1: '10',
      $E$1: '20',
      $A$2: '=SUM(A1:E1)'
    }
  })
})
