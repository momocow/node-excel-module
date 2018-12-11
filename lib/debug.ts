import { join, resolve as resolvePath, dirname } from 'path'
import { writeFile } from 'fs'
import * as mkdirp from 'mkdirp'

import libDebug from 'debug'
export const debug = libDebug('excel-module')

export function dumpToFile (file: string, str: string): Promise<string> {
  if (debug.enabled) {
    return new Promise<string>(function (resolve, reject) {
      require('find-pkg')(__dirname)
        .then(function (pkgDir: string) {
          const DUMP_FILE = resolvePath(join(dirname(pkgDir), '.debug'), file)
          mkdirp(dirname(DUMP_FILE), (err) => {
            if (err) return reject(err)
            writeFile(DUMP_FILE, str, 'utf8', (err) => {
              if (err) return reject(err)
              resolve(DUMP_FILE)
            })
          })
        })
        .catch((err: Error) => reject(err))
    })
  }

  return Promise.resolve('')
}
