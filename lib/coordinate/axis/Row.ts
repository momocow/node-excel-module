import Axis from './Axis'
import Index from './Index'
import ValueError from '../../errors/ValueError'

export default class Row extends Axis {
  public readonly index: Index

  constructor (index: Index) {
    super()
    this.index = index
  }

  get label () {
    return `$${this.index.base1}`
  }

  static from (label: string): Row {
    const index = Number(label.replace(/^\$/, ''))
    if (Number.isNaN(index)) throw new ValueError(`The index should not be NaN.`)
    // index is 1-based index
    return new Row(new Index(index, 1))
  }
}
