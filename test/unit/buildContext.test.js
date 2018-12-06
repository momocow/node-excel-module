const test = require('ava')

const { buildContext } = require('../../dist/lib/utils')

test('buildContext()', async t => {
  t.plan(1)

  const table = {
    A1: '1',
    B1: '2',
    C1: '=SUM(D1,E1)',
    D1: '10',
    E1: '20',
    F1: '###',
    G1: '***',
    A2: '=SUM(A1:E1)'
  }

  const context = await buildContext([ '1!A2' ], (label, done) => {
    done(1, label, table[label])
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
