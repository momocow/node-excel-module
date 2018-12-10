import Axis from './Axis'
import Index from './Index'
import ValueError from '../../errors/ValueError'

export default class Sheet extends Axis {
  public readonly index: Index

  constructor (index: number) {
    if (Number.isNaN(index)) throw new ValueError(`The index should not be NaN.`)

    super()

    this.index = new Index(index, 0)
  }

  get label () {
    return `${this.index}!`
  }

  get abs () {
    return true
  }

  static from (label: string): Sheet {
    return new Sheet(Number(label.replace(/!$/, '')))
  }
}
