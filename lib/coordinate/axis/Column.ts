import * as base26 from 'base26'

import Axis from './Axis'
import Index from './Index'
import ValueError from '../../errors/ValueError'

export default class Column extends Axis {
  public readonly index: Index

  constructor (index: Index) {
    super()

    this.index = index
  }

  get label () {
    return `$${base26.to(this.index.base1).toUpperCase()}`
  }

  static from (label: string): Column {
    const index = label.replace(/^\$/, '').toLowerCase()
    // base26 is 1-based index
    return new Column(new Index(base26.from(index), 1))
  }
}
