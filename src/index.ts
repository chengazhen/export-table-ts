import * as XLSX from 'xlsx'
import type { WorkBook, WritingOptions, WorkSheet, Range } from 'xlsx'
import { saveAs } from 'file-saver'
type CellStyle = {
  alignment: {
    horizontal: string
    vertical: string
  }
  [index: string]: any
}

type HeadProp = {
  name: string
  prop?: string
  exeFun?: Function
  child?: HeadProp[]
  summable?: unknown
}

type PropType = {
  [index: string]: unknown
  child?: PropType[]
}

type DataProp = PropType[]

class ExprotExcel {
  defaultCellStyle: CellStyle = {
    alignment: {
      horizontal: 'center',
      vertical: 'center'
    }
  }

  constructor(defaultCellStyle?: CellStyle) {
    defaultCellStyle ? (this.defaultCellStyle = defaultCellStyle) : ''
  }

  exportData(
    tableData: DataProp,
    tableHeader: HeadProp[],
    sheetName: string = 'table'
  ) {
    // excel表头
    const excelHeader: Array<string | number>[] = this.buildHeader(tableHeader)

    const headerRows: number = excelHeader.length

    const dataList: Array<string | number>[] = this.extractData(
      tableData,
      tableHeader
    )

    excelHeader.push(...dataList, [])

    const merges: Range[] = this.doMerges(excelHeader)

    const ws: WorkSheet = this.aoa_to_sheet(excelHeader, headerRows)

    ws['!merges'] = merges

    ws['!freeze'] = {
      xSplit: '1',
      ySplit: '' + headerRows,
      topLeftCell: 'B' + (headerRows + 1),
      activePane: 'bottomRight',
      state: 'frozen'
    }

    // 列宽
    ws['!cols'] = [{ wpx: 165 }]

    let workbook: WorkBook = {
      SheetNames: [sheetName],
      Sheets: {}
    }

    workbook.Sheets[sheetName] = ws
    // excel样式
    const wopts: WritingOptions = {
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary',
      cellStyles: true
    }

    // 表格样式
    const wbout = XLSX.write(workbook, wopts)

    let blob: Blob = new Blob([this.s2ab(wbout)], {
      type: 'application/octet-stream'
    })

    saveAs(blob, sheetName + '.xlsx')
  }

  private s2ab(s: string) {
    let buf: ArrayBuffer = new ArrayBuffer(s.length)
    let view: Uint8Array = new Uint8Array(buf)
    for (let i = 0; i !== s.length; ++i) {
      view[i] = s.charCodeAt(i) & 0xff
    }
    return buf
  }

  private aoa_to_sheet(datas: Array<string | number>[], headerRows: number) {
    const ws: WorkSheet = {}
    const range: { s: { c: number; r: number }; e: { c: number; r: number } } =
      {
        s: { c: 10000000, r: 10000000 },
        e: { c: 0, r: 0 }
      }

    datas.forEach((data, R) => {
      data.forEach((item, C) => {
        if (range.s.r > R) {
          range.s.r = R
        }
        if (range.s.c > C) {
          range.s.c = C
        }
        if (range.e.r < R) {
          range.e.r = R
        }
        if (range.e.c < C) {
          range.e.c = C
        }

        const cell: {
          v: number | string
          s: Record<string, unknown>
          t?: string
        } = {
          v: item || '',

          s: Object.assign({}, this.defaultCellStyle) //这里是因为防止浅复制导致数据地址指向同一数据
        }

        // 头部列表加边框
        if (R < headerRows) {
          const style = Object.assign(cell.s, {
            border: {
              top: { style: 'thin', color: { rgb: '000000' } },
              left: { style: 'thin', color: { rgb: '000000' } },
              bottom: { style: 'thin', color: { rgb: '000000' } },
              right: { style: 'thin', color: { rgb: '000000' } }
            },
            fill: {
              patternType: 'solid',
              fgColor: { theme: 3, tint: 0.3999755851924192, rgb: 'F5F7FA' },
              bgColor: { theme: 7, tint: 0.3999755851924192, rgb: 'F5F7FA' }
            }
          })
          cell.s = style
        }

        const cell_ref = XLSX.utils.encode_cell({ c: C, r: R })

        if (typeof cell.v === 'number') {
          cell.t = 'n'
        } else if (typeof cell.v === 'boolean') {
          cell.t = 'b'
        } else {
          cell.t = 's'
        }

        ws[cell_ref] = cell
      })
    })

    if (range.s.c < 10000000) {
      ws['!ref'] = XLSX.utils.encode_range(range)
    }
    return ws
  }

  /**
   * @description: 需要合并的数据
   * @param {*} arr
   * @return {*}
   */
  private doMerges(arrs: Array<string | number>[]) {
    const merges: Range[] = []
    arrs.forEach((arr, y) => {
      let colSpan: number = 0
      arr.forEach((row, x) => {
        if (row === '!$COL_SPAN_PLACEHOLDER') {
          arr[x] = ''
          if (x + 1 === arr.length) {
            merges.push({ s: { r: y, c: x - colSpan - 1 }, e: { r: y, c: x } })
          }
          colSpan++
        } else if (colSpan > 0 && x > colSpan) {
          merges.push({
            s: { r: y, c: x - colSpan - 1 },
            e: { r: y, c: x - 1 }
          })
          colSpan = 0
        } else {
          colSpan = 0
        }
      })
    })

    const deep = arrs.length
    arrs[0].forEach((element, x) => {
      let rowSpan: number = 0
      for (let y = 0; y < deep; y++) {
        if (arrs[y][x] === '!$ROW_SPAN_PLACEHOLDER') {
          arrs[y][x] = ''
          if (y + 1 === deep) {
            merges.push({ s: { r: y - rowSpan, c: x }, e: { r: y, c: x } })
          }
          rowSpan++
        } else if (rowSpan > 0 && y > rowSpan) {
          merges.push({
            s: { r: y - rowSpan - 1, c: x },
            e: { r: y - 1, c: x }
          })
          rowSpan = 0
        } else {
          rowSpan = 0
        }
      }
    })

    return merges
  }

  /**
   * 根据选中的数据和展示的列，生成结果
   * @param selectionData
   * @param revealList
   */
  private extractData(selectionData: DataProp, revealList: HeadProp[]) {
    // 列
    let headerList = this.flat(revealList, Infinity)
    // 导出的结果集
    let excelRows: Array<string | number>[] = []
    // 如果有child集合的话会用到
    let dataKeys: Set<string> = new Set(Object.keys(selectionData[0]))

    selectionData.some(e => {
      if (e.child && e.child.length > 0) {
        let childKeys: string[] = Object.keys(e.child[0])
        for (let i = 0; i < childKeys.length; i++) {
          dataKeys.delete(childKeys[i])
        }
        return true
      }
    })

    this.flatData(selectionData, (list: DataProp) => {
      excelRows.push(...this.buildExcelRow(dataKeys, headerList, list))
    })

    return excelRows
  }

  private buildExcelRow(
    mainKeys: Set<string>,
    headers: HeadProp[],
    rawDataList: DataProp
  ) {
    const sumCols: Array<number> = []
    const rows: Array<string | number>[] = []
    let index: number = 0
    for (const rawData of rawDataList) {
      const cols: Array<string | number> = []
      for (const header of headers) {
        if (
          header.prop &&
          rawData['rowSpan'] === 0 &&
          mainKeys.has(header.prop)
        ) {
          cols.push('!$ROW_SPAN_PLACEHOLDER')
        } else {
          let value: string | number = ''
          if (header.exeFun && typeof header.exeFun === 'function') {
            value = header.exeFun(rawData)
          } else {
            if (header.prop) {
              value = rawData[header.prop] as string
            }
          }

          cols.push(value)

          if (header['summable'] && typeof value === 'number') {
            sumCols[index] = (sumCols[index] ? sumCols[index] : 0) + value
          }
        }

        index++
      }

      rows.push(cols)
    }

    if (sumCols.length > 0) {
      rows.push(...this.sumRowHandle())
    }

    return rows
  }

  private sumRowHandle() {
    return []
  }

  /**
   * @description: 展开多维头部
   * @param {HeadProp[]} revealList 需要展开的数组
   * @return { HeadProp[]}
   */
  // 扁平头部
  private flat(revealList: HeadProp[], d: number = 1): HeadProp[] {
    const result: HeadProp[] = []
    return d > 0
      ? revealList.reduce(
          (acc, val) =>
            acc.concat(
              Array.isArray(val.child) ? this.flat(val.child, d - 1) : val
            ),
          result
        )
      : revealList.slice()
  }

  /**
   * @description:铺平数组
   * @param {*} list 表格数据
   * @param {*} eachDataCallBack 构建的行数据
   * @return {*}
   */
  private flatData(list: DataProp, eachDataCallBack: Function) {
    let resultList = []
    for (let i = 0; i < list.length; i++) {
      let data: PropType = list[i]
      let rawDataList: DataProp = []
      // 每个子元素都和父元素合并成一条数据
      if (data.child && data.child.length > 0) {
        for (let j = 0; j < data.child.length; j++) {
          delete data.child[j].bsm
          let copy = Object.assign({}, data, data.child[j])
          rawDataList.push(copy)
          copy['rowSpan'] = j > 0 ? 0 : data.child.length
        }
      } else {
        data['rowSpan'] = 1
        rawDataList.push(data)
      }
      resultList.push(...rawDataList)

      eachDataCallBack(rawDataList)
    }
    return resultList
  }

  // flatData2(revealList: DataProp, eachDataCallBack): Record<string, unknown>[] {
  //   for (const item of revealList) {
  //     const result: PropType[] = []

  //     item.child && item.child
  //   }
  //   // return d > 0
  //   //   ? revealList.reduce(
  //   //       (acc, val) =>
  //   //         acc.concat(
  //   //           Array.isArray(val.child) ? this.flatData2(val.child, d - 1) : val
  //   //         ),
  //   //       result
  //   //     )
  //   //   : revealList.slice()
  // }

  /**
   * 构建excel表头
   * @param revealList 列表页面展示的表头
   * @returns {[]} excel表格展示的表头
   */
  private buildHeader(revealList: HeadProp[]) {
    let excelHeader: Array<string[]> = []
    // 构建生成excel表头需要的数据结构
    this.getHeader(revealList, excelHeader, 0, 0)
    // 多行表头长短不一，短的向长的看齐，不够的补上行合并占位符
    let max = Math.max(...excelHeader.map(a => a.length))
    excelHeader
      .filter(e => e.length < max)
      .forEach(e => this.pushRowSpanPlaceHolder(e, max - e.length))
    return excelHeader
  }

  /**
   * 生成头部
   * @param headers 展示的头部
   * @param excelHeader excel头部
   * @param deep 深度
   * @param perOffset 前置偏移量
   * @returns {number}  后置偏移量
   */
  private getHeader(
    headers: HeadProp[],
    excelHeader: Array<string[]>,
    deep: number,
    perOffset: number
  ): number {
    let offset: number = 0
    let cur: string[] = excelHeader[deep]

    if (!cur) {
      cur = excelHeader[deep] = []
    }

    this.pushRowSpanPlaceHolder(cur, perOffset - cur.length)
    for (let i = 0; i < headers.length; i++) {
      let head = headers[i]
      cur.push(head.name)
      if (
        head.hasOwnProperty('child') &&
        Array.isArray(head.child) &&
        head.child.length > 0
      ) {
        let childOffset = this.getHeader(
          head.child,
          excelHeader,
          deep + 1,
          cur.length - 1
        )
        // 填充列合并占位符
        this.pushColSpanPlaceHolder(cur, childOffset - 1)
        offset += childOffset
      } else {
        offset++
      }
    }
    return offset
  }

  /**
   * @description: 填充列合并占位符
   * @param {string} arr
   * @param {number} count
   * @return {*}
   */
  private pushColSpanPlaceHolder(arr: string[], count: number) {
    for (let i = 0; i < count; i++) {
      arr.push('!$COL_SPAN_PLACEHOLDER')
    }
  }
  /**
   * 填充行合并占位符
   * */
  private pushRowSpanPlaceHolder(arr: any[], count: number) {
    for (let i = 0; i < count; i++) {
      arr.push('!$ROW_SPAN_PLACEHOLDER')
    }
  }
}

const exportExcel = new ExprotExcel()

const exportData = exportExcel.exportData.bind(exportExcel)

export default ExprotExcel

export { exportData }
