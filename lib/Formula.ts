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
  return Reference.from(`${cell.sheet || defaultSheet.index.base1}!${cell.column.label}${cell.row.label}`)
}

export class Formula {
  constructor (
    public formula: string,
    public sheet: Sheet
  ) {
  }

  getMembers (lookUp: (cell: Reference, done: (err: Error | null, result: any) => void) => void): Reference[] {
    debug('Formula#getOperandCells()')

    let errs: Error[] = []
    let cells: Reference[] = []
    const parser = new FormulaParser()
    parser.on('callCellValue', (cell: FormulaParserReference, done: Function) => {
      try {
        debug(
          'Formula#getOperandCells(): event "callCellValue" [%s!%s%s]',
          cell.sheet,
          cell.column.label,
          cell.row.label
        )
        const ref = createReference(cell, this.sheet)
        debug('Formula#getOperandCells():   => %o', ref.toString())
        cells.push(ref)

        lookUp(ref, (err, result) => {
          if (err) {
            errs.push(err)
            return
          }

          done(result)
        })

      } catch (e) {
        errs.push(e)
        throw e
      }
    })
    parser.on('callRangeValue', (start: FormulaParserReference, end: FormulaParserReference, done: Function) => {
      try {
        debug('Formula#getOperandCells(): event "callRangeValue" [%s!%s%s] => [%s!%s%s]',
          start.sheet,
          start.column.label,
          start.row.label,
          end.sheet,
          end.column.label,
          end.row.label
        )
        const table = new Range(
          createReference(start, this.sheet),
          createReference(end, this.sheet)
        ).toArray()
        debug('Formula#getOperandCells():   => %o', table.map(ref => ref.toString()))
        cells = cells.concat(table.reduce((acc, row) => {
          acc.push(...row)
          return acc
        }, []))

        done(table.map(row => row.map(ref => {
          let v
          lookUp(ref, (err, result) => {
            if (err) {
              errs.push(err)
              return
            }
            v = result
          })
          return v
        })))
      } catch (e) {
        errs.push(e)
        throw e
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
