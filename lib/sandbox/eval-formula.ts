import { FormulaParserReference } from '../formula-parser/FormulaParserReference'

interface EvalContext extends Record<string, any> {
  args?: string[]
  entry?: string
}

/**
 * Context always has the coordinates in absolute format.
 */
export function evalFormula (context: EvalContext, ...evalArgs: any[]) {

  // /***** Common Methods *****/

  function createError (name: string, msg?: string): Error {
    const e = new Error(msg)
    e.name = name
    return e
  }

  /**
   * @see https://github.com/jameslaneconkling/base26
   */
  function toBase26 (decimal: number): string {
    let out = ''

    while (true) {
      out = String.fromCharCode(((decimal) % 26) + 97) + out
      decimal = Math.floor((decimal) / 26)

      if (decimal === 0) {
        break
      }
    }

    return out
  }

  function parseRange (
    start: FormulaParserReference,
    end: FormulaParserReference
  ): string[] {
    const ret: string[] = []
    for (let y = start.row.index; y <= end.row.index; y++) {
      for (let x = start.column.index; x <= end.column.index; x++) {
        ret.push(`${toBase26(x)}${y + 1}`)
      }
    }

    return ret
  }

  // /**************************/

  // context replacement
  if (context.args) {
    context.args.forEach((arg, i) => {
      context[arg] = evalArgs[i]
    })
  }

  const FormulaParser = require('hot-formula-parser').Parser
  const parser = new FormulaParser()

  parser.on('callCellValue', function ({ label, sheet }: FormulaParserReference, done: Function) {
    done(parser.parse(context[`${sheet}!${label}`].replace(/^=/, '')).result)
  })

  parser.on('callRangeValue', function (
    start: FormulaParserReference,
    end: FormulaParserReference,
    done: Function
  ) {
    done(parseRange(start, end).map(
      (label: string) => parser.parse(context[label].replace(/^=/, '')).result
    ))
  })

  if (!context.entry) {
    throw createError('EntryError', 'The entry cell is not defined.')
  }

  const { error, result } = parser.parse(context.entry.replace(/^=/, ''))
  if (error) {
    throw createError('FormulaError', error)
  }
  return result
}
