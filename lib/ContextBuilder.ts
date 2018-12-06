import _merge = require('lodash/merge')

import { debug } from './debug'

import { Coordinate } from './coordinate'

import {
  toAbsCoord,
  parseRange
} from './utils'

const { Parser: FormulaParser } = require('hot-formula-parser')

export class BuilderNotInitializedError extends Error {
  constructor () {
    super(`The context builder has not been completely initialized yet.`)
    this.name = 'BuilderNotInitializedError'
  }
}

export class ContextBuilder {
  public readonly context: Record<string, Record<string, any>> = {}

  private sheet: number = 1

  public isInitialized (): boolean {
    return true
  }

  /**
   * Change sheet
   */
  private chsheet (sheet: number) {
    this.sheet = Math.max(1, sheet)
  }

  private setContext (coord: Coordinate, value: any) {
    if (!this.context[coord.sheet]) this.context[coord.sheet] = {}
    this.context[coord.sheet][coord.label] = value
  }

  public async build (entries: string[]): Promise<Record<string, any>> {
    if (!this.isInitialized()) throw new BuilderNotInitializedError()

    return this.context
  }
}

export async function buildContext (
  entries: string[],
  lookUp: (label: string, done: (sheet: number, label: string, value: any) => void) => any,
  context: Record<string, any> = {}
): Promise<Record<string, any>> {
  debug('buildContext(): entry point = %o', entries)
  debug('buildContext(): context = %o', context)

  const formulaParser = new FormulaParser()

  for (let cell of entries) {
    await new Promise(function (resolve, reject) {
      lookUp(cell, async (sheet, label, value) => {
        label = toAbsCoord(label)

        debug('lookUp(): %o', { sheet, label, value })

        try {
          _merge(context, {
            [sheet]: {
              [label]: value
            }
          })

          if (typeof value === 'string' && value.startsWith('=')) {
            const materials: string[] = []
            // @TODO Cross sheets references have not been supported yet
            formulaParser.on('callCellValue', function ({ label }: Coordinate) {
              try {
                debug('buildContext(): event "callCellValue" [label=%o]', label)
                materials.push(label)
              } catch (e) {
                return reject(e)
              }
            })
            formulaParser.on('callRangeValue', function (start: Coordinate, end: Coordinate) {
              try {
                debug('buildContext(): event "callRangeValue" [start=%o, end=%o]', start.label, end.label)
                materials.push(...parseRange(start, end))
              } catch (e) {
                return reject(e)
              }
            })
            formulaParser.parse(value.slice(1))
            formulaParser.off('callCellValue')
            formulaParser.off('callRangeValue')

            debug('buildContext(): related cells = %o', materials)
            if (materials.length > 0) {
              _merge(context, await buildContext(materials, lookUp, context))
              debug('buildContext(): new context = %o', context)
            }
          }
        } catch (err) {
          return reject(err)
        }

        resolve()
      })
    })

  }

  return context
}
