import _merge = require('lodash/merge')

import { debug } from './debug'

import { Coordinate } from './coordinate'

import {
  toAbsCoord,
  parseRange
} from './utils'

const { Parser: FormulaParser } = require('hot-formula-parser')

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
