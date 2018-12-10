import ValueError from '../../errors/ValueError'

export default class Index {
  private _base0: number

  constructor (index: number, base: number = 0) {
    if (index < base) {
      throw new ValueError(`The index (${index}) should not be less than the base (${base}).`)
    }
    this._base0 = index - base
  }

  public get base0 () {
    return this.toBase(0)
  }

  public get base1 () {
    return this.toBase(1)
  }

  public toBase (base: number) {
    return this._base0 + base
  }

  public valueOf () {
    return this._base0
  }

  public toJSON () {
    return this._base0
  }

  public toString () {
    return '' + this._base0
  }
}
