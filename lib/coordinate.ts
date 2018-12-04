import * as base26 from 'base26'

abstract class Axis {
  public abstract index: number
  public abstract label: string
  public abstract abs: boolean

  public valueOf () {
    return this.index
  }

  public toString () {
    return this.label
  }
}

export class Row extends Axis {
  public label: string
  public index: number
  public abs: boolean
  constructor (label: string | number) {
    super()
    this.index = Number(label)
    this.label = String(label)
    this.abs = typeof label === 'string' ? label.startsWith('$') : true
  }
}

export class Column extends Axis {
  public label: string
  public index: number = 0
  public abs: boolean

  constructor (labelOrIndex: string | number) {
    super()

    if (typeof labelOrIndex === 'string') {
      this.label = labelOrIndex.toUpperCase()
      this.index = base26.from(this.label)
      this.abs = labelOrIndex.startsWith('$')
    } else {
      this.label = base26.to(labelOrIndex).toUpperCase()
      this.index = labelOrIndex
      this.abs = true
    }
  }
}

export class InvalidCoordError extends Error {
  constructor (text: string) {
    super(`Invalid coordinate is given: "${text}".`)
    this.name = 'InvalidCoordError'
  }
}

export class Coordinate {
  public sheet: number = 1
  public label: string
  public column: Column
  public row: Row

  constructor (col: Column, row: Row, sheet: number)
  constructor (label: string)
  constructor (arg1: string | Column, row?: Row, sheet: number = 1) {
    if (typeof arg1 === 'string') {
      let matched = arg1.match(/(?:(\d+)!)?\$?([A-Z]+)\$?(\d+)/)
      if (!matched) throw new InvalidCoordError(arg1)

      this.label = arg1
      this.sheet = Number(matched[1])
      this.column = new Column(matched[2])
      this.row = new Row(matched[3])
    } else {
      this.column = arg1
      this.row = row
      this.sheet = sheet
      this.label = 
    }
  }

  isAbsolute () {
    return this.column.abs && this.row.abs
  }
}

export class RelativeCoord {
  private offset: { x: number, y: number } = { x: 0, y: 0 }
  constructor (
    public rel: Coordinate,
    public pivot: Coordinate
  ) {
    if (!this.pivot.isAbsolute()) throw new TypeError(`The pivot should be absolute.`)

    this.offset = {
      x: this.rel.column.index - this.pivot.column.index,
      y: this.rel.row.index - this.pivot.row.index
    }
  }

  public resolveFrom (target: Coordinate) {
    if (!target.isAbsolute()) throw new TypeError(`The target should be absolute.`)

    target
  }
}
