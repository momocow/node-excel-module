import { FormulaParserAxis } from './FormulaParserAxis'

export interface FormulaParserReference {
  row: FormulaParserAxis
  column: FormulaParserAxis
  label: string
  sheet: string
}
