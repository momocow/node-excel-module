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
