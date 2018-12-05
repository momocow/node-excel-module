import * as serialize from 'serialize-javascript'

import { debug } from './debug'
import {
  CoordRelation,
  Coordinate
} from './coordinate'

const { Parser: FormulaParser } = require('hot-formula-parser')

export function wrapFunc (
  body: Function,
  context: Record<string, any> = {},
  contextKey: string = 'context'
): Function {
  return Function(`
    const ${contextKey} = ${serialize(context)};
    return (${body.toString()})(${contextKey}, ...Array.from(arguments));
    `
  )
}

export function parseRange (start: Coordinate, end: Coordinate) {
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
      formulaParser.on('callCellValue', function ({ label }: Coordinate) {
        debug('buildContext(): event "callCellValue" [label=%o]', label)
        materials.push(label)
      })
      formulaParser.on('callRangeValue', function (start: Coordinate, end: Coordinate) {
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

export function resolveSharedFormula (formula: string, from: Coordinate, to: Coordinate) {
  debug('resolveSharedFormula(): formula = %s', formula)
  debug('resolveSharedFormula(): from = %s', from.label)
  debug('resolveSharedFormula(): to = %s', to.label)

  let lastIndex = 0
  let ret = ''

  function replaceLabel (label: string) {
    debug('replaceLabel(): label = "%s"', label)
    const relCoord = new CoordRelation(new Coordinate(label), from)
    const labelIndex = formula.indexOf(label)
    debug('replaceLabel(): label index = %d', labelIndex)
    if (labelIndex >= 0) {
      ret += formula.substring(lastIndex, labelIndex) + relCoord.resolveFrom(to).label
      lastIndex = labelIndex + label.length
    }
  }

  const formulaParser = new FormulaParser()
  formulaParser.on('callCellValue', function ({ label }: Coordinate) {
    try {
      replaceLabel(label)
    } catch (e) {
      console.error(e)
    }
  })
  formulaParser.on('callRangeValue', function (start: Coordinate, end: Coordinate) {
    replaceLabel(start.label)
    replaceLabel(end.label)
  })
  formulaParser.parse(formula)

  if (lastIndex < formula.length) {
    ret += formula.substr(lastIndex)
  }

  debug('resolveSharedFormula(): resolved = %s', ret)

  return ret
}
