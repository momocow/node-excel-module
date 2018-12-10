export default class CrossSheetRangeError extends Error {
  constructor () {
    super(`Cross-sheet range is not allowed.`)
  }
}
