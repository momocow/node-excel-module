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

    if (typeof label === 'string') {
      this.abs = label.startsWith('$')
      this.index = Number(label.replace(/^\$/, '')) - 1
      this.label = label
    } else {
      this.abs = true
      this.index = label
      this.label = '$' + (this.index + 1)
    }
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
      this.abs = labelOrIndex.startsWith('$')
      this.index = base26.from(this.label.replace(/^\$/, '').toLowerCase()) - 1
    } else {
      this.label = '$' + base26.to(labelOrIndex + 1).toUpperCase()
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

  constructor (col: Column, row: Row, sheet?: number)
  constructor (label: string)
  constructor (arg1: string | Column, row?: Row, sheet: number = 1) {
    if (typeof arg1 === 'string') {
      let matched = arg1.match(/(?:(\d+)!)?(\$?[A-Z]+)(\$?\d+)/)
      if (!matched) throw new InvalidCoordError(arg1)

      this.sheet = Number(matched[1] || 1)
      this.column = new Column(matched[2])
      this.row = new Row(matched[3])
    } else {
      this.column = arg1
      this.row = row as Row
      this.sheet = sheet
    }
    this.label = this.column.toString() + this.row.toString()
  }

  shift (deltaCol: number = 0, deltaRow: number = 0): Coordinate {
    return new Coordinate(
      new Column(this.column.index + deltaCol),
      new Row(this.row.index + deltaRow),
      this.sheet
    )
  }

  isAbsolute () {
    return this.column.abs && this.row.abs
  }

  valueOf () {
    return [
      this.column.index,
      this.row.index
    ]
  }

  toJSON () {
    return {
      label: this.label,
      sheet: this.sheet
    }
  }

  toString () {
    return this.sheet + '!' + this.label
  }
}

export class CoordRelation {
  private offset: [ number, number ] = [ 0, 0 ]
  constructor (
    public rel: Coordinate,
    public pivot: Coordinate
  ) {
    if (!this.pivot.isAbsolute()) throw new TypeError(`The pivot should be absolute.`)

    this.offset = [
      this.rel.column.index - this.pivot.column.index,
      this.rel.row.index - this.pivot.row.index
    ]
  }

  public resolveFrom (target: Coordinate): Coordinate {
    if (!target.isAbsolute()) throw new TypeError(`The target should be absolute.`)
    return target.shift(this.offset[0], this.offset[1])
  }
}
