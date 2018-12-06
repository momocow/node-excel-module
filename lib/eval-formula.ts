export function evalFormula (context: Record<string, any>, ...evalArgs: any[]) {
  function _get (obj: Record<string, any>, path: string, defaultValue?: any): any {
    for (let tok of path.split('.')) {
      if (typeof obj === 'object' && tok in obj) {
        obj = obj[tok]
      } else {
        return defaultValue
      }
    }
    return obj
  }

  function _set (obj: Record<string, any>, path: string, value: any) {
    const arrPath = path.split('.')
    for (let tok of arrPath.slice(0, -1)) {
      if (typeof obj[tok] !== 'object') {
        obj[tok] = {}
      }
      obj = obj[tok]
    }

    obj[arrPath[arrPath.length - 1]] = value
  }

  function parseCoord (label: string): string {
    const matched = label.match(/(\d+!)?\$?([A-Z]+)\$?(\d+)/)
    if (!matched) {
      let e = new Error('Invalid cell named "' + label + '".')
      e.name = 'InvalidCellError'
      throw e
    }
    return (matched[1] || 1) + '.$' + matched[2] + '$' + matched[3]
  }

  // context replacement
  if (Array.isArray(context.args)) {
    context.args.forEach((arg, i) => {
      arg = parseCoord(arg)
      _set(context, arg, evalArgs[i] || _get(context, arg))
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
  const { error, result } = parser.parse(_get(context, context.entry).replace(/^=/, ''))
  if (error) {
    const e = new Error(error)
    e.name = 'FormulaError'
    throw e
  }
  return result
}
