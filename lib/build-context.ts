import { debug } from './debug'
import Reference from './coordinate/Reference'
import { Formula } from './Formula'

export default async function buildContext (
  entries: Reference[],
  lookUp: (label: Reference, done: (err: Error | null, value: any) => void) => any,
  context: Record<string, any> = {}
): Promise<Record<string, any>> {
  debug('buildContext(): entry point = %o', entries)
  debug('buildContext(): initial context = %o', context)

  for (let cell of entries) {
    await new Promise(function (resolve, reject) {
      lookUp(cell, async (err, value) => {
        if (err) return reject(err)
        debug('buildContext() -> lookUp: cell "%s" = %o', cell.label, value)
        try {
          context[cell.toString()] = value

          if (typeof value === 'string' && Formula.isFormula(value)) {
            debug('buildContext(): formula %o', value)

            const formula = Formula.from(value, cell.sheet)
            const operands = formula.getOperandCells()

            debug('buildContext(): operand cells = %o', operands)
            if (operands.length > 0) {
              Object.assign(context, await buildContext(operands, lookUp, context))
              debug('buildContext(): new context = %o', context)
            }
          }
        } catch (err) {
          return reject(err)
        }
        resolve()
      })
    })

  }

  return context
}
