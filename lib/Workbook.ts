import {
  Workbook as WorkbookBase,
  ValueType,
  CellErrorValue,
  CellHyperlinkValue,
  CellRichTextValue,
  CellFormulaValue,
  FormulaType,
  CellSharedFormulaValue
} from 'exceljs'

import {
  wrapFunc,
  parseRange,
  buildContext,
  formatValue,
  resolveSharedFormula
} from './utils'

import {
  Coordinate
} from './coordinate'

import { evalFormula } from './eval-formula'

import { debug } from './debug'

interface CompileOptions {
}
Number(1)
type CellType = |
  FunctionConstructor |
  NumberConstructor |
  StringConstructor |
  BooleanConstructor

interface CellSpec {
  cell: string
  type: CellType
  args?: string[]
}

class EmptyWorkbookError extends Error {
  constructor () {
    super(`The workbook is empty.`)
    this.name = 'EmptyWorkbookError'
  }
}

export default class Workbook extends WorkbookBase {
  public compile<T extends Record<string, CellSpec>> (
    spec: T,
    options?: CompileOptions
  ): {
    [K in keyof T]: ReturnType<T[K]['type']>
  } {
    if (this.worksheets.length === 0) throw new EmptyWorkbookError()

    debug('# Spec %o', spec)

    /**
     * MUST be serializable
     */
    const context: Record<string, any> = buildContext(
      Object.values(spec).map(c => c.cell),
      (label: string) => {
        let coord = new Coordinate(label)

        debug('## Label[%s]: %o', label, coord.toJSON())

        const worksheet = this.getWorksheet(coord.sheet)
        let curCell = worksheet.getCell(coord.label)
        coord = new Coordinate(curCell.$col$row)

        // Ensure the cell to be the master
        if (curCell.type === ValueType.Merge) {
          debug('##   Use master cell')
          curCell = curCell.master
          coord = new Coordinate(curCell.$col$row)
        }

        debug('##   = %o', curCell.value)

        switch (curCell.type) {
          // primitive values
          case ValueType.Date:
          case ValueType.Boolean:
          case ValueType.Null:
          case ValueType.Number:
          case ValueType.String:
            return curCell.value
          // text values
          case ValueType.Hyperlink:
            return (curCell.value as CellHyperlinkValue).text
          case ValueType.RichText:
            return (curCell.value as CellRichTextValue).richText
              .map(t => t.text)
              .join('')
          // @TODO
          // SharedString is not documented (https://github.com/guyonroche/exceljs#value-types)
          // Leave it unimplemented until any reproducible scenarios
          // case ValueType.SharedString:
          case ValueType.Error:
            return (curCell.value as CellErrorValue).error
          case ValueType.Formula:
            if (curCell.formulaType === FormulaType.Shared) {
              curCell = worksheet.getCell((curCell.value as CellSharedFormulaValue).sharedFormula)
              return '=' + resolveSharedFormula(
                (curCell.value as CellFormulaValue).formula,
                new Coordinate(curCell.$col$row),
                coord
              )
            }
            return '=' + (curCell.value as CellFormulaValue).formula
        }
      }
    )
    debug('# Context %o', context)

    const exports: Record<string, any> = {}

    for (let key of Object.keys(spec)) {
      debug('## Exporting member[%s]', key)

      let cellType: CellType = spec[key].type ? spec[key].type
        : spec[key].args ? Function
          : String
      let cell: string = spec[key].cell

      debug('### Exported as a %s', cellType.name)
      debug('### Cell[%s]', cell)
      debug('###   = %o', context[cell])

      if (cellType === Function) {
        exports[key] = typeof context[cell] === 'string' && context[cell].startsWith('=')
          ? wrapFunc(evalFormula, {
            entry: cell,
            ...context,
            parseRange,
            args: spec[key].args
          }, 'localContext')
        : wrapFunc(function () {
          return context.value
        }, {
          value: context[cell]
        })
      } else {
        exports[key] = exports[cell] = formatValue(
          context[cell], (cellType as Exclude<CellType, FunctionConstructor>))
      }
    }

    debug('# Exports %o', exports)

    return exports as {
      [K in keyof T]: ReturnType<T[K]['type']>
    }
  }
}
