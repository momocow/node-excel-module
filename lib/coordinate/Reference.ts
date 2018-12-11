import Column from './axis/Column'
import Row from './axis/Row'
import Sheet from './axis/Sheet'
import ValueError from '../errors/ValueError'
import Vector from './Vector'
import Labelable from './Labelable'
import Index from './axis/Index'

/**
 * Reference is the absolute cell reference used in excel
 */
export default class Reference implements Labelable {
  constructor (
    public readonly sheet: Sheet,
    public readonly column: Column,
    public readonly row: Row
  ) {
  }

  public offset (vec: Vector) {
    return new Reference(
      this.sheet,
      new Column(new Index(this.column.index.base0 + vec.colOffset, 0)),
      new Row(new Index(this.row.index.base0 + vec.rowOffset, 0))
    )
  }

  public equals (another: Reference): boolean {
    return this.column.equals(another.column) &&
      this.row.equals(another.row) &&
      this.sheet.equals(another.sheet)
  }

  toString () {
    return `${this.sheet.toString()}${this.column.toString()}${this.row.toString()}`
  }

  get label () {
    return `${this.column.toString()}${this.row.toString()}`
  }

  /**
   * @param label Omitting "$" is allowed but every cell ref is treated as an absolute ref.
   */
  static from (label: string): Reference {
    const matched = label.match(/(\d+)!\$?([a-zA-Z]+)\$?(\d+)/)
    if (!matched) throw new ValueError(`The ref label ("${label}") is invalid.`)

    return new Reference(
      Sheet.from(matched[1]),
      Column.from(matched[2]),
      Row.from(matched[3])
    )
  }
}
