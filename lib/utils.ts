import * as serialize from 'serialize-javascript'

export function wrapFunc (
  body: Function,
  context: Record<string, any> = {}
): Function {
  // tslint:disable-next-line:no-eval
  return eval(`
    (
      function () {${
          '\n  ' +
          Object.entries(context)
            .map(([ k, v ]) => `const ${k} = ${serialize(v, { space: 4 })};`)
            .join('\n' + ' '.repeat(2))
        }
        return (${body.toString()})(${Object.keys(context).join(', ')}, ...Array.from(arguments));
      }
    )
  `) as Function
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
