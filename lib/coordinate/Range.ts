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

  toArray () {
    const unitXVector = new Vector(1, 0)
    const unitYVector = new Vector(0, 1)
    const ret = []

    let curRef = this.start
    for (;
      curRef.row.index.base0 <= this.end.row.index.base0;
      curRef = curRef.offset(unitYVector)
    ) {
      for (;
        curRef.column.index.base0 <= this.end.column.index.base0;
        curRef = curRef.offset(unitXVector)
      ) {
        ret.push(curRef)
      }

      // because curRef.column.index.base0 > this.end.column.index.base0
      // so the for loop right above breaks
      // curRef should be reset to the same x (col) as this.end
      curRef = curRef.offset(new Vector(-1, 0))
    }

    return ret
  }

  get label () {
    return this.start.label + ':' + this.end.column.label + this.end.row.label
  }

  static from (start: string, end: string) {
    return new Range(Reference.from(start), Reference.from(end))
  }
}
