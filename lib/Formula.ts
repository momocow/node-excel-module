import Reference from './coordinate/Reference'
import Range from './coordinate/Range'
import Sheet from './coordinate/axis/Sheet'

import { debug } from './debug'
const { Parser: FormulaParser } = require('hot-formula-parser')

interface FormulaParserAxis {
  index: number
  label: string
  isAbsolute: boolean
}

interface FormulaParserReference {
  row: FormulaParserAxis
  column: FormulaParserAxis
  label: string
  sheet: string
}

function createReference (cell: FormulaParserReference, defaultSheet: Sheet): Reference {
  return Reference.from(`${cell.sheet || defaultSheet.index.base1}!${cell.label}`)
}

export class Formula {
  constructor (
    public formula: string,
    public sheet: Sheet
  ) {
  }

  getOperandCells (): Reference[] {
    debug('Formula#getOperandCells()')

    let errs: Error[] = []
    let cells: Reference[] = []
    const parser = new FormulaParser()
    parser.on('callCellValue', (cell: FormulaParserReference) => {
      try {
        debug('Formula#getOperandCells(): event "callCellValue" [%s!%s]', cell.sheet, cell.label)
        const ref = createReference(cell, this.sheet)
        debug('Formula#getOperandCells():   => %o', ref.toString())
        cells.push(ref)
      } catch (e) {
        errs.push(e)
      }
    })
    parser.on('callRangeValue', (start: FormulaParserReference, end: FormulaParserReference) => {
      try {
        debug('Formula#getOperandCells(): event "callRangeValue" [%s!%s] => [%s!%s]',
          start.sheet, start.label, end.sheet, end.label)
        const refs = new Range(
          createReference(start, this.sheet),
          createReference(end, this.sheet)
        ).toArray()
        debug('Formula#getOperandCells():   => %o', refs.map(ref => ref.toString()))
        cells = cells.concat(refs)
      } catch (e) {
        errs.push(e)
      }
    })
    parser.parse(this.formula)

    if (errs.length > 0) throw errs[0]

    return cells
  }

  valueOf () {
    return this.formula
  }

  toString () {
    return `=${this.formula}`
  }

  static from (formula: string, sheet: string | Sheet): Formula {
    return new Formula(
      formula.replace(/^=/, ''),
      typeof sheet === 'string'
        ? Sheet.from(sheet)
        : sheet
    )
  }

  static isFormula (literal: string): boolean {
    return literal.startsWith('=')
  }
}
