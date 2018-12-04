import * as serialize from 'serialize-javascript'

import { debug } from './debug'
import { WorkbookCoord } from './coordinate'

const { Parser: FormulaParser } = require('hot-formula-parser')

export function wrapFunc (
  body: Function,
  context: Record<string, any> = {}
): Function {
  return Function(`
    const context = ${serialize(context)};
    return (${body.toString()})(context, ...Array.from(arguments));
    `
  )
}

export function parseRange (start: WorkbookCoord, end: WorkbookCoord) {
  const ret: string[] = []
  for (let y = start.row.index; y <= end.row.index; y++) {
    for (
      let x = start.column.label;
      x.charCodeAt(0) <= end.column.label.charCodeAt(0);
      x = String.fromCharCode(x.charCodeAt(0) + 1)
    ) {
      ret.push(`${x}${y + 1}`)
    }
  }

  return ret
}

export function offsetByCharCode (char: string, offset: number): string {
  return String.fromCharCode(char.charCodeAt(0) + offset)
}

export function buildContext (
  entries: string[],
  lookUp: (label: string) => any,
  context: Record<string, any> = {}
): Record<string, any> {
  debug('buildContext(): entry point = %o', entries)
  debug('buildContext(): context = %o', context)

  const formulaParser = new FormulaParser()

  for (let cell of entries) {
    context[cell] = lookUp(cell)

    if (typeof context[cell] === 'string' && context[cell].startsWith('=')) {
      const materials: string[] = []
      // @TODO Cross sheets references have not been supported yet
      formulaParser.on('callCellValue', function ({ label }: WorkbookCoord) {
        debug('buildContext(): event "callCellValue" [label=%o]', label)
        materials.push(label)
      })
      formulaParser.on('callRangeValue', function (start: WorkbookCoord, end: WorkbookCoord) {
        debug('buildContext(): event "callRangeValue" [start=%o, end=%o]', start.label, end.label)
        materials.push(...parseRange(start, end))
      })
      formulaParser.parse(context[cell].slice(1))
      formulaParser.off('callCellValue')
      formulaParser.off('callRangeValue')

      debug('buildContext(): related cells = %o', materials)
      if (materials.length > 0) {
        Object.assign(context, buildContext(materials, lookUp, context))
        debug('buildContext(): new context = %o', context)
      }
    }
  }

  return context
}

type ValueType = StringConstructor | NumberConstructor | BooleanConstructor

export function formatValue (
  value: any,
  type: ValueType
): ReturnType<ValueType> {
  switch (type) {
    case Number:
      return parseFloat(value)
    case String:
      return `${value}`
    case Boolean:
      return Boolean(value)
    default:
      throw new Error(`Unknown type, "${type.name}".`)
  }
}
