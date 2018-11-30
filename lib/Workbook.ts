import {
  Workbook as WorkbookBase,
  ValueType,
  CellErrorValue,
  CellHyperlinkValue,
  CellRichTextValue,
  CellFormulaValue
} from 'exceljs'

const { Parser: FormulaParser } = require('hot-formula-parser')
const formulaParser = new FormulaParser()

interface CompileOptions {
}

interface CellSpec {
  cell: string
  args?: Record<string, string>
}

type APIFunction = (...args: any[]) => boolean | number | string
type APIExport = null | boolean | number | string | APIFunction

interface WorkbookCoordinate {
  sheet: string | number,
  label: string
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
  private buildFormulaAST (formula: string) {
    formulaParser.parse(formula)
  }

  public compile<T extends Record<string, CellSpec>> (
    spec: T,
    options?: CompileOptions
  ): Record<keyof T, APIExport> {
    if (this.worksheets.length === 0) throw new EmptyWorkbookError()

    const exports: Record<string, any> = {}

    for (let fn of Object.keys(spec)) {
      const { cell, args } = spec[fn]
      const coord = Workbook.parseCoord(cell)

      if (!coord) throw new UnrecognizedCellError(cell)

      const worksheet = this.getWorksheet(coord.sheet)
      let curCell = worksheet.getCell(coord.label)

      // Ensure the cell to be the master
      if (curCell.type === ValueType.Merge) {
        curCell = curCell.master
      }

      switch (curCell.type) {
        // primitive values
        case ValueType.Date:
        case ValueType.Boolean:
        case ValueType.Null:
        case ValueType.Number:
        case ValueType.String:
          exports[fn] = exports[coord.label] = curCell.value
          break
        // text values
        case ValueType.Hyperlink:
          exports[fn] = exports[coord.label] = (curCell.value as CellHyperlinkValue).text
          break
        case ValueType.RichText:
          exports[fn] = exports[coord.label] = (curCell.value as CellRichTextValue).richText.map(tok => tok.text).join('')
          break
        // @TODO
        // SharedString is not documented (https://github.com/guyonroche/exceljs#value-types)
        // Leave it unimplemented until any reproducible scenarios
        // case ValueType.SharedString:
        case ValueType.Error:
          exports[fn] = (curCell.value as CellErrorValue).error
          break
        case ValueType.Formula:
          this.buildFormulaAST((curCell.value as CellFormulaValue).formula)
      }
    }

    return Object.keys(spec).reduce((acc, k) => {
      acc[k] = () => ''
      return acc
    }, {} as Record<keyof T, APIExport>)
  }

  static parseCoord (str: string): WorkbookCoordinate | null {
    let matched = str.match(/(?:([\s\w-_]+)!)?([A-Z]+)([1-9]+[0-9]*)/)
    if (matched) {
      return {
        sheet: matched[1] || 1,
        label: matched[2]
      }
    }

    return null
  }
}
