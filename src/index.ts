import xlsx from 'xlsx'

type CellStyle = {
  alignment: {
    horizontal: string
    vertical: string
  }
}

type HeadProp = {
  name: string
  prop?: string
  child?: HeadProp[]
}

type PropType = {
  [index: string]: unknown
  child?: PropType[]
}

type DataProp = PropType[]

class ExprotExcel {
  defaultCellStyle: CellStyle
  constructor() {
    this.defaultCellStyle = {
      alignment: {
        horizontal: 'center',
        vertical: 'center'
      }
    }
  }

  exportData(
    tableData: Record<string, unknown>[],
    tableHeader: HeadProp[],
    sheetName: string = 'table'
  ) {
    // excel表头
    const excelHeader: string[][] = this.buildHeader(tableHeader)

    const headerRows: number = excelHeader.length

    let dataList = this.extractData(tableData, tableHeader)
  }

  /**
   * 根据选中的数据和展示的列，生成结果
   * @param selectionData
   * @param revealList
   */
  private extractData(selectionData: HeadProp[], revealList: HeadProp[]) {
    // 列
    let headerList = this.flat(revealList)
    // 导出的结果集
    let excelRows: string[] = []
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

    this.flatData(selectionData, list => {
      excelRows.push(...this.buildExcelRow(dataKeys, headerList, list))
    })
    return excelRows
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
  flatData(list: DataProp, eachDataCallBack) {
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
      if (typeof eachDataCallBack === 'function') {
        eachDataCallBack(rawDataList)
      }
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
  pushColSpanPlaceHolder(arr: string[], count: number) {
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
