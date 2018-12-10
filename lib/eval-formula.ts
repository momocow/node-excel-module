interface NormalizedLabel {
  label: string
  sheet: number
  column: string
  row: number
}

interface EvalContext extends Record<string, any> {
  args?: string[]
  entry?: string
}

export function evalFormula (context: EvalContext, ...evalArgs: any[]) {

  // /***** Common Methods *****/

  function createError (name: string, msg?: string): Error {
    const e = new Error(msg)
    e.name = name
    return e
  }

  function normalizeLabel (label: string): NormalizedLabel {
    const sheetEnsured = label.indexOf('!') <= 0 ? '1!' + label.replace(/^!/, '') : label
    const matched = sheetEnsured.match(/(\d+)!\$?([a-zA-Z]+)\$?(\d+)/)
    if (!matched) throw createError('InvalidCellLabelError', `Cannot recognize "${label}".`)
    return {
      label: sheetEnsured,
      sheet: Number(matched[1]),
      column: String(matched[2]),
      row: Number(matched[3])
    }
  }

  // /**************************/

  // context replacement
  if (context.args) {
    context.args.forEach((arg, i) => {
      const { label: argLabel } = normalizeLabel(arg)
      context[argLabel] = evalArgs[i]
    })
  }

  const FormulaParser = require('hot-formula-parser').Parser
  const parser = new FormulaParser()

  parser.on('callCellValue', function ({ label }: { label: string }, done: Function) {
    done(
      parser.parse(
        String(_get(context, parseCoord(label)))
          .replace(/^=/, '')
      ).result
    )
  })

  parser.on('callRangeValue', function (start: any, end: any, done: Function) {
    done(
      context.parseRange(start, end)
        .map((label: string) => _get(context, parseCoord(label)))
        .map((f: string) => parser.parse(String(f).replace(/^=/, '')).result)
    )
  })

  const { error, result } = parser.parse(_get(context, ).replace(/^=/, ''))
  if (error) {
    const e = new Error(error)
    e.name = 'FormulaError'
    throw e
  }
  return result
}
