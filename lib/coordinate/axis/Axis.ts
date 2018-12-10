import Index from './Index'
import Labelable from '../Labelable'

export default abstract class Axis implements Labelable {
  public readonly abstract index: Index
  public readonly abstract label: string

  public valueOf () {
    return this.index.valueOf()
  }

  public toString () {
    return this.label
  }

  public toJSON () {
    return {
      index: this.index.valueOf(),
      label: this.label
    }
  }

  public equals (another: Axis): boolean {
    return this.label === another.label
  }
}
