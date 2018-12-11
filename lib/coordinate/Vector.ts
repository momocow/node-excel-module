import Reference from './Reference'

/**
 * Vector is the relative cell reference used in excel
 */
export default class Vector {
  constructor (public colOffset: number, public rowOffset: number) {
  }

  public normalize () {
    return new Vector(this.colOffset / this.length, this.rowOffset / this.length)
  }

  get length () {
    return Math.sqrt(this.colOffset * this.colOffset + this.rowOffset * this.rowOffset)
  }

  static from (label: string, pivot: string): Vector {
    const targetRef = Reference.from(label)
    const pivotRef = Reference.from(pivot)
    return new Vector(
      targetRef.column.index.base0 - pivotRef.column.index.base0,
      targetRef.row.index.base0 - pivotRef.row.index.base0
    )
  }
}
