import Axis from './Axis'
import Index from './Index'
import ValueError from '../../errors/ValueError'

export default class Row extends Axis {
  public readonly index: Index

  constructor (index: number) {
    if (Number.isNaN(index)) throw new ValueError(`The index should not be NaN.`)

    super()
    this.index = new Index(index, 0)
  }

  get label () {
    return `$${this.index.base1}`
  }

  static from (label: string): Row {
    const abs = label.startsWith('$')
    const index = Number(label.replace(/^\$/, ''))
    // index is 1-based index
    return new Row(index - 1)
  }
}
