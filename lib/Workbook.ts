import {
  Workbook as WorkbookBase,
  ValueType,
  CellErrorValue,
  CellHyperlinkValue,
  CellRichTextValue
} from 'exceljs'

import buildContext from './build-context'

import {
  wrapFunc,
  formatValue,
  mapSheetName2Index,
  ensureSheetNumber
} from './utils'

import { evalFormula } from './sandbox/eval-formula'

import { debug } from './debug'
import Reference from './coordinate/Reference'

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

    const ensureSheet = (label: string) => ensureSheetNumber(
      mapSheetName2Index(label, (sheet) => this.getWorksheet(sheet).id)
    )

    /**
     * MUST be serializable
     */
    const context: Record<string, any> = await buildContext(
      Object.values(spec).map(c => Reference.from(ensureSheet(c.cell))),
      (cell: Reference, done: (err: Error | null, value: any) => void) => {
        debug('## Cell[%s]', cell)

        const worksheet = this.getWorksheet(cell.sheet.index.base1)
        let curCell = worksheet.getCell(cell.label)

        // Ensure the cell to be the master
        if (curCell.type === ValueType.Merge) {
          debug('##   Use master Cell[%s]', curCell.$col$row)
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
            return done(null, curCell.value)
          // text values
          case ValueType.Hyperlink:
            return done(null, (curCell.value as CellHyperlinkValue).text)
          case ValueType.RichText:
            return done(
              null,
              (curCell.value as CellRichTextValue).richText
                .map(t => t.text)
                .join('')
            )
          // @TODO
          // SharedString is not documented (https://github.com/guyonroche/exceljs#value-types)
          // Leave it unimplemented until any reproducible scenarios
          // case ValueType.SharedString:
          case ValueType.Error:
            return done(null, (curCell.value as CellErrorValue).error)
          case ValueType.Formula:
            return done(null, '=' + curCell.formula)
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
      let cell: Reference = Reference.from(ensureSheet(spec[key].cell))

      debug('### Exported as a %s', cellType.name)
      debug('### Cell[%s]', cell)

      const cellValue = context[cell.toString()]
      debug('###   = %o', cellValue)

      if (cellType === Function) {
        exports[key] = typeof cellValue === 'string' && cellValue.startsWith('=')
          ? wrapFunc(evalFormula, {
            entry: cell.toString().replace('!', '.'),
            ...context,
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
