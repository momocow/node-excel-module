import Axis from './Axis'
import Index from './Index'
import ValueError from '../../errors/ValueError'

export default class Sheet extends Axis {
  public readonly index: Index

  constructor (index: Index) {
    super()

    this.index = index
  }

  get label () {
    return `${this.index}!`
  }

  get abs () {
    return true
  }

  static from (label: string): Sheet {
    const index = Number(label.replace(/!$/, ''))
    if (Number.isNaN(index)) throw new ValueError(`The index should not be NaN.`)
    return new Sheet(new Index(index, 1))
  }
}
