import { Workbook } from 'exceljs'
import { handleStyle } from './style'


type TableJson = {
  data: {
    name: string
    cells: ({ v?: string; s?: number } | null)[][]
    plain?: string[][]
    cols: { width: number }[]
    rows: { height: number }[]
    default_row_height?: number
    merged?: { from: { column: number; row: number }; to: { column: number; row: number } }[]
  }[]
  styles: Record<string, string>[]
}

const HEIGHT_COEF = 0.75;
const WIDTH_COEF = 8.5;

function getCellCoords(r: number, col: number) {
  let c = ''
  while (col != 0) {
    let residual = col % 26
    col = Math.floor(col / 26)
    if (residual == 0) {
      residual += 26
      col -= 1
    }
    c = String.fromCharCode(residual + 64) + c
  }
  return `${c}${r}`
}

export async function import_to_xlsx(rawData: TableJson) {
  const workbook = new Workbook()
  rawData.styles = rawData.styles.map((s) => handleStyle(s))
  rawData.data.forEach((s) => {
    const sheet = workbook.addWorksheet(s.name, {
      properties: {
        defaultRowHeight: typeof s.default_row_height !== 'undefined' ? s.default_row_height * HEIGHT_COEF : 15.75
      }
    })

    s.cols.forEach((c, i) => {
      if (c) {
        sheet.getColumn(i + 1).width = c.width / WIDTH_COEF
      }
    })

    s.rows.forEach((r, i) => {
      if (r) {
        sheet.getRow(i + 1).height = r.height * HEIGHT_COEF
      }
    })

    s.cells.forEach((r, i) => {
      r.forEach((c, j) => {
        if (c) {
          const cell = sheet.getCell(getCellCoords(i + 1, j +1))
          if (!isNaN(c.s!)) {
            cell.style = rawData.styles[c.s!]
          }
          cell.value = c.v
        }
      })
    })

    if (s.merged) {
      s.merged.forEach((m) => {
        sheet.mergeCells(getCellCoords(m.from.row + 1, m.from.column + 1), getCellCoords(m.to.row + 1, m.to.column + 1))
      })
    }
  })

  return await workbook.xlsx.writeBuffer()
}