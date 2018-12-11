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
  ): string[][] {
    const rety: string[][] = []
    for (let y = start.row.index; y <= end.row.index; y++) {
      const retx: string[] = []
      for (let x = start.column.index; x <= end.column.index; x++) {
        retx.push(`${start.sheet}!$${toBase26(x).toUpperCase()}$${y + 1}`)
      }
      rety.push(retx)
    }

    return rety
  }

  // /**************************/

  // context replacement
  if (context.args) {
    context.args.forEach((arg, i) => {
      context[arg] = evalArgs[i] === undefined ? context[arg] : evalArgs[i]
    })
  }

  const errors: Error[] = []
  const FormulaParser = require('hot-formula-parser').Parser
  const parser = new FormulaParser()

  parser.on('callCellValue', function ({ label, sheet }: FormulaParserReference, done: Function) {
    try {
      done(parser.parse(`${context[`${sheet}!${label}`]}`.replace(/^=/, '')).result)
    } catch (e) {
      errors.push(e)
    }
  })

  parser.on('callRangeValue', function (
    start: FormulaParserReference,
    end: FormulaParserReference,
    done: Function
  ) {
    try {
      done(parseRange(start, end).map(
        (labels: string[]) => labels.map(label => parser.parse(`${context[label]}`.replace(/^=/, '')).result)
      ))
    } catch (e) {
      errors.push(e)
    }
  })

  if (!context.entry) {
    throw createError('EntryError', 'The entry cell is not defined.')
  }

  const { error, result } = parser.parse(`${context[context.entry]}`.replace(/^=/, ''))

  if (errors.length > 0) throw errors[0]

  if (error) {
    throw createError('FormulaError', error)
  }
  return result
}
