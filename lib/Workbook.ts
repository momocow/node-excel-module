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
  resolveSharedFormula,
  toAbsCoord
} from './utils'

import {
  Coordinate
} from './coordinate'

import { evalFormula } from './eval-formula'

import { debug } from './debug'

interface CompileOptions {
  exposeCells?: boolean
}

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
  public async compile<T extends Record<string, CellSpec>> (
    spec: T,
    options: CompileOptions = {}
  ): Promise<{
    [K in keyof T]: ReturnType<T[K]['type']>
  }> {
    if (this.worksheets.length === 0) throw new EmptyWorkbookError()

    debug('# Spec %o', spec)

    /**
     * MUST be serializable
     */
    const context: Record<string, any> = await buildContext(
      Object.values(spec).map(c => c.cell),
      (label: string, done: (sheet: number, label: string, value: any) => void) => {
        let coord = new Coordinate(label)
        debug('## Label[%s]: %o', label, coord.toJSON())

        const worksheet = this.getWorksheet(coord.sheet)
        let curCell = worksheet.getCell(coord.label)

        // fix coordinate to be absolute
        coord = new Coordinate(curCell.$col$row)

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
            return done(coord.sheet, coord.label, curCell.value)
          // text values
          case ValueType.Hyperlink:
            return done(coord.sheet, coord.label, (curCell.value as CellHyperlinkValue).text)
          case ValueType.RichText:
            return done(
              coord.sheet,
              coord.label,
              (curCell.value as CellRichTextValue).richText
                .map(t => t.text)
                .join('')
            )
          // @TODO
          // SharedString is not documented (https://github.com/guyonroche/exceljs#value-types)
          // Leave it unimplemented until any reproducible scenarios
          // case ValueType.SharedString:
          case ValueType.Error:
            return done(coord.sheet, coord.label, (curCell.value as CellErrorValue).error)
          case ValueType.Formula:
            if (curCell.formulaType === FormulaType.Shared) {
              curCell = worksheet.getCell((curCell.value as CellSharedFormulaValue).sharedFormula)
              return done(
                coord.sheet,
                coord.label,
                '=' + resolveSharedFormula(
                  (curCell.value as CellFormulaValue).formula,
                  new Coordinate(curCell.$col$row),
                  coord
                )
              )
            }
            return done(coord.sheet, coord.label, '=' + (curCell.value as CellFormulaValue).formula)
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
      let cell: Coordinate = new Coordinate(toAbsCoord(spec[key].cell))

      debug('### Exported as a %s', cellType.name)
      debug('### Cell[%s]', cell.toString())

      const cellValue = context[cell.sheet][cell.label]
      debug('###   = %o', cellValue)

      if (cellType === Function) {
        exports[key] = typeof cellValue === 'string' && cellValue.startsWith('=')
          ? wrapFunc(evalFormula, {
            entry: cell.toString().replace('!', '.'),
            ...context,
            parseRange,
            args: spec[key].args
          }, 'localContext')
        : wrapFunc(function () {
          return context.value
        }, {
          value: cellValue
        })

        if (options.exposeCells) {
          exports[cell.toString()] = exports[key]
        }
      } else {
        exports[key] = formatValue(
          cellValue, (cellType as Exclude<CellType, FunctionConstructor>))
        if (options.exposeCells) {
          exports[cell.toString()] = exports[key]
        }
      }
    }

    debug('# Exports %o', exports)

    return exports as {
      [K in keyof T]: ReturnType<T[K]['type']>
    }
  }
}
