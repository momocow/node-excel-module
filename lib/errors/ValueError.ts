export default class ValueError extends Error {
  constructor (msg: string = 'invalid value') {
    super(msg)
    this.name = 'ValueError'
  }
}
