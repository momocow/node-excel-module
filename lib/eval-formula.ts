export function evalFormula (context: Record<string, any>, ...evalArgs: any[]) {
  // context replacement
  if (Array.isArray(context.args)) {
    context.args.forEach((arg, i) => {
      context[arg] = evalArgs[i] || context[arg]
    })
  }

  const FormulaParser = require('hot-formula-parser').Parser
  const parser = new FormulaParser()
  parser.on('callCellValue', function ({ label }: { label: string }, done: Function) {
    label = label.replace(/\$(\w+)/g, '$1')
    done(parser.parse(String(context[label]).replace(/^=/, '')).result)
  })
  parser.on('callRangeValue', function (start: any, end: any, done: Function) {
    done(
      context.parseRange(start, end).map((label: string) => context[label])
        .map((f: string) => parser.parse(String(f).replace(/^=/, '')).result)
    )
  })
  const { error, result } = parser.parse(context[context.entry].replace(/^=/, ''))
  if (error) {
    const e = new Error(error)
    e.name = 'FormulaError'
    throw e
  }
  return result
}
