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
  formatValue
} from './utils'

import { evalFormula } from './sandbox/eval-formula'

import { debug, dumpToFile } from './debug'
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
  private normalizeCoords (label: string, defaultSheet: number): string {
    return label.replace(
      /(?:'?([^']+)'?!)?\$?([a-zA-Z]+)\$?(\d+)(?::\$?([a-zA-Z]+)\$?(\d+))?/g,
      (
        label,
        sheet: string | undefined,
        col1: string,
        row1: string,
        col2?: string,
        row2?: string
      ): string => {
        debug('normalizeCoords(): from %o', { label, sheet, col1, row1, col2, row2 })

        const _sheet = `${sheet ? this.getWorksheet(sheet).id : defaultSheet}!`
        const _ref1 = `$${col1}$${row1}`
        const _ref2 = col2 && row2 ? `:$${col2}$${row2}` : ''
        debug('normalizeCoords():   => to "%s"', _sheet + _ref1 + _ref2)

        return _sheet + _ref1 + _ref2
      }
    )
  }

  public async compile<T extends Record<string, CellSpec>> (
    spec: T,
    options: CompileOptions = {}
  ): Promise<{
    [K in keyof T]: ReturnType<T[K]['type']>
  }> {
    if (this.worksheets.length === 0) throw new EmptyWorkbookError()

    debug('# Spec => %s', await dumpToFile(
      'spec.json',
      JSON.stringify(
        spec,
        (k, v) => typeof v === 'function' ? v.toString() : v,
        2
      )
    ))

    /**
     * MUST be serializable
     */
    const context: Record<string, any> = await buildContext(
      Object.values(spec).map(c => Reference.from(this.normalizeCoords(c.cell, 1))),
      (cell: Reference, done: (err: Error | null, value: any) => void) => {
        debug('## Cell[%s]', cell)

        const worksheet = this.getWorksheet(cell.sheet.index.base1)
        let curCell = worksheet.getCell(cell.label)

        debug('##   = %o', curCell.value)

        switch (curCell.type) {
          // merge cell is empty
          case ValueType.Merge:
            return done(null, '')
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
            return done(
              null,
              '=' + this.normalizeCoords(curCell.formula, cell.sheet.index.base1)
            )
        }
      }
    )
    debug('# Context => %s', await dumpToFile('context.json', JSON.stringify(context, null, 2)))

    const exports: Record<string, any> = {}

    for (let key of Object.keys(spec)) {
      debug('## Exporting member[%s]', key)

      let cellType: CellType = spec[key].type ? spec[key].type
        : spec[key].args ? Function
          : String
      let cell: Reference = Reference.from(this.normalizeCoords(spec[key].cell, 1))

      debug('### Exported as a %s', cellType.name)
      debug('### Cell[%s]', cell)

      const cellValue = context[cell.toString()]
      debug('###   = %o', cellValue)

      if (cellType === Function) {
        exports[key] = typeof cellValue === 'string' && cellValue.startsWith('=')
          ? wrapFunc(evalFormula, {
            entry: cell.toString(),
            ...context,
            args: !spec[key].args
              ? undefined
              // spec[key].args! to suppress TS2532
              // (why "!spec[key].args" does not persuade the typeguard?)
              : spec[key].args!.map(
                argCell => Reference.from(this.normalizeCoords(argCell, 1)).toString()
              )
          })
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

    const dataExports: Record<string, any> = {}
    for (let k of Object.keys(exports)) {
      if (typeof exports[k] === 'function') {
        debug('# Exports[%s] => %s', k, await dumpToFile(`exports/${k}.js`, exports[k].toString()))
      } else {
        dataExports[k] = exports[k]
      }
    }

    return exports as {
      [K in keyof T]: ReturnType<T[K]['type']>
    }
  }
}
