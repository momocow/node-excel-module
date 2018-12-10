import * as serialize from 'serialize-javascript'

export function wrapFunc (
  body: Function,
  context: Record<string, any> = {},
  contextKey: string = 'context'
): Function {
  return Function(`
    const ${contextKey} = ${serialize(context)};
    return (${body.toString()})(${contextKey}, ...Array.from(arguments));
    `
  )
}

type ValueType = StringConstructor | NumberConstructor | BooleanConstructor

export function formatValue (
  value: any,
  type: ValueType
): ReturnType<ValueType> {
  switch (type) {
    case Number:
      return parseFloat(value)
    case String:
      return `${value}`
    case Boolean:
      return Boolean(value)
    default:
      throw new Error(`Unknown type, "${type.name}".`)
  }
}

export function mapSheetName2Index (
  label: string,
  lookUp: (sheet: string) => number
): string {
  return label.replace(/^.*?!/, function (matched) {
    return lookUp(matched.substr(0, matched.length - 1)) + '!'
  })
}

export function ensureSheetNumber (label: string): string {
  if (label.indexOf('!') <= 0) label = '1!' + label.replace(/^!/, '')
  return label
}
