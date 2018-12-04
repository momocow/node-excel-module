const test = require('ava')

const { buildContext } = require('../../dist/lib/utils')

test('buildContext()', t => {
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

  const context = buildContext([ 'A2' ], (label) => {
    return table[label]
  })

  t.deepEqual(context, {
    A1: '1',
    B1: '2',
    C1: '=SUM(D1,E1)',
    D1: '10',
    E1: '20',
    A2: '=SUM(A1:E1)'
  })
})
