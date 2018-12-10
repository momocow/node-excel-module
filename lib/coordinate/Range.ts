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
    const unitVector = Vector.from(this.end.label, this.start.label).normalize()
    const ret = []
    for (
      let curRef = this.start;
      !curRef.equals(this.end);
      curRef = curRef.offset(unitVector)
    ) {
      ret.push(curRef)
    }
    ret.push(this.end)

    return ret
  }

  get label () {
    return this.start.label + ':' + this.end.column.label + this.end.row.label
  }

  static from (start: string, end: string) {
    return new Range(Reference.from(start), Reference.from(end))
  }
}
