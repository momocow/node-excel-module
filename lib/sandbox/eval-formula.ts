import { FormulaParserReference } from '../formula-parser/FormulaParserReference'

const context: Record<string, any> = {}

/**
 * Context always has the coordinates in absolute format.
 */
export function evalFormula (entry: string, args: string[], ...evalArgs: any[]) {
  const FormulaParser = require('hot-formula-parser').Parser
  const parser = new FormulaParser()

  // /***** Common Methods *****/

  function parse (expr: any) {
    if (typeof expr === 'string' && expr.startsWith('=')) {
      const { error, result } = parser.parse(expr.replace(/^=/, ''))
      if (error) {
        throw createError('FormulaError', error)
      }
      return result
    } else {
      return expr
    }
  }

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

  const localContext: Record<string, any> = Object.assign({}, context)

  // context replacement
  if (Array.isArray(args)) {
    args.forEach((arg, i) => {
      localContext[arg] = evalArgs[i] === undefined ? context[arg] : evalArgs[i]
    })
  }

  const errors: Error[] = []

  parser.on('callCellValue', function (
    { row: { label: rowLabel }, column: { label: colLabel }, sheet }: FormulaParserReference,
    done: Function
  ) {
    try {
      done(parse(localContext[`${sheet}!$${colLabel}$${rowLabel}`]))
    } catch (e) {
      errors.push(e)
      throw e
    }
  })

  parser.on('callRangeValue', function (
    start: FormulaParserReference,
    end: FormulaParserReference,
    done: Function
  ) {
    try {
      done(parseRange(start, end).map(
        (labels: string[]) => labels.map(label => parse(localContext[label]))
      ))
    } catch (e) {
      errors.push(e)
      throw e
    }
  })

  if (!entry) {
    throw createError('EntryError', 'The entry cell is not defined.')
  }

  const result = parse(localContext[entry])

  if (errors.length > 0) throw errors[0]
  return result
}
