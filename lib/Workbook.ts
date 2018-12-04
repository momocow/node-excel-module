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
  formatValue
} from './utils'

import {
  parseCoord,
  WorkbookCoord
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

class UnrecognizedCellError extends Error {
  constructor (cell: string) {
    super(`Unrecognized cell: "${cell}".`)
    this.name = 'UnrecognizedCellError'
  }
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
        const coord = parseCoord(label) as Pick<WorkbookCoord, 'sheet' | 'label'>
        if (!coord) throw new UnrecognizedCellError(label)

        debug('## Label[%s]: %o', label, coord)

        const worksheet = this.getWorksheet(coord.sheet)
        let curCell = worksheet.getCell(coord.label)

        // Ensure the cell to be the master
        if (curCell.type === ValueType.Merge) {
          debug('##   Use master cell')
          curCell = curCell.master
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
          })
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
