import Labelable from './coordinate/Labelable'
import Reference from './coordinate/Reference'
import Range from './coordinate/Range'
import Sheet from './coordinate/axis/Sheet'
import Vector from './coordinate/Vector'

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
    let cells: Reference[] = []
    const parser = new FormulaParser()
    parser.on('callCellValue', (cell: FormulaParserReference) => {
      cells.push(createReference(cell, this.sheet))
    })
    parser.on('callRangeValue', (start: FormulaParserReference, end: FormulaParserReference) => {
      cells = cells.concat(
        new Range(
          createReference(start, this.sheet),
          createReference(end, this.sheet)
        ).toArray()
      )
    })
    parser.parse(this.formula)
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
