import { Readable, Writable } from 'stream'

import Workbook from './lib/Workbook'

export async function from (src: string | Readable): Promise<Workbook> {
  const workbook = new Workbook()
  if (src instanceof Readable) {
    return new Promise<Workbook>(function (resolve, reject) {
      src.pipe(workbook.xlsx.createInputStream() as Writable)
        .on('finish', () => resolve(workbook))
        .on('error', err => reject(err))
    })
  }
  await workbook.xlsx.readFile(src)
  return workbook
}
