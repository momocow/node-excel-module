import Labelable from './Labelable'
import Reference from './Reference'
import CrossSheetRangeError from '../errors/CrossSheetRangeError'
import Vector from './Vector'

export default class Range implements Labelable {
  constructor (
    public start: Reference,
    public end: Reference
  ) {
    if (!this.start.sheet.equals(this.end.sheet)) throw new CrossSheetRangeError()
  }

  toArray (): Reference[][] {
    const unitXVector = new Vector(1, 0)
    const unitYVector = new Vector(0, 1)
    const table: Reference[][] = []

    let curRef = this.start
    for (;
      curRef.row.index.base0 <= this.end.row.index.base0;
      curRef = curRef.offset(unitYVector)
    ) {
      const row: Reference[] = []

      for (;
        curRef.column.index.base0 <= this.end.column.index.base0;
        curRef = curRef.offset(unitXVector)
      ) {
        row.push(curRef)
      }

      // because curRef.column.index.base0 > this.end.column.index.base0
      // so the for loop right above breaks
      // curRef should be reset in x-axis
      curRef = curRef.offset(
        new Vector(-1 * (this.end.column.index.base0 - this.start.column.index.base0 + 1), 0)
      )

      table.push(row)
    }

    return table
  }

  get label () {
    return this.start.label + ':' + this.end.column.label + this.end.row.label
  }

  static from (start: string, end: string) {
    return new Range(Reference.from(start), Reference.from(end))
  }
}
