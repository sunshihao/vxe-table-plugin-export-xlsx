/* eslint-disable no-unused-vars */
import XEUtils from 'xe-utils/methods/xe-utils'
import {
  VXETable,
  Table,
  InterceptorExportParams,
  InterceptorImportParams,
  ColumnConfig,
  ExportOptons
} from 'vxe-table/lib/vxe-table'
import XLSX from 'xlsx'
/* eslint-enable no-unused-vars */

let _vxetable: typeof VXETable

function getCellLabel (column: ColumnConfig, cellValue: any) {
  if (cellValue) {
    switch (column.cellType) {
      case 'string':
        return XEUtils.toString(cellValue)
      case 'number':
        if (!isNaN(cellValue)) {
          return Number(cellValue)
        }
        break
      default:
        if (cellValue.length < 12 && !isNaN(cellValue)) {
          return Number(cellValue)
        }
        break
    }
  }
  return cellValue
}

function getFooterCellValue ($table: Table, opts: ExportOptons, rows: any[], column: ColumnConfig) {
  var cellValue = XEUtils.toString(rows[$table.$getColumnIndex(column)])
  return cellValue
}

function toBuffer (wbout: any) {
  let buf = new ArrayBuffer(wbout.length)
  let view = new Uint8Array(buf)
  for (let index = 0; index !== wbout.length; ++index) view[index] = wbout.charCodeAt(index) & 0xFF
  return buf
}

function exportXLSX (params: InterceptorExportParams) {
  const { $table, options, columns, datas } = params
  const { sheetName, isHeader, isFooter, original, message, footerFilterMethod } = options
  const colHead: { [key: string]: any } = {}
  const footList: { [key: string]: any }[] = []
  // const rowList = datas

  if (isHeader) {
    columns.forEach((column) => {
      colHead[column.id] = XEUtils.toString(original ? column.property : column.getTitle())
    })
  }

  // 新增部分
  const rowList = datas.map(item => {
    const rest: any = {}
    columns.forEach((column) => {
      rest[column.id] = getCellLabel(column, item[column.id])
    })
    return rest
  })

  if (isFooter) {
    const { footerData } = $table.getTableData()
    const footers = footerFilterMethod ? footerData.filter(footerFilterMethod) : footerData
    footers.forEach((rows) => {
      const item: { [key: string]: any } = {}
      columns.forEach((column) => {
        item[column.id] = getFooterCellValue($table, options, rows, column)
      })
      footList.push(item)
    })
  }
  const book = XLSX.utils.book_new()
  const sheet = XLSX.utils.json_to_sheet((isHeader ? [colHead] : []).concat(rowList).concat(footList), { skipHeader: true })
  // 转换数据
  XLSX.utils.book_append_sheet(book, sheet, sheetName)
  const wbout = XLSX.write(book, { bookType: 'xlsx', bookSST: false, type: 'binary' })
  const blob = new Blob([toBuffer(wbout)], { type: 'application/octet-stream' })
  // 保存导出
  downloadFile(blob, options)
  if (message !== false) {
    _vxetable.modal.message({ message: _vxetable.t('vxe.table.expSuccess'), status: 'success' })
  }
}

function downloadFile (blob: Blob, options: ExportOptons) {
  if (window.Blob) {
    const { filename, type } = options
    if (navigator.msSaveBlob) {
      navigator.msSaveBlob(blob, `${filename}.${type}`)
    } else {
      var linkElem = document.createElement('a')
      linkElem.target = '_blank'
      linkElem.download = `${filename}.${type}`
      linkElem.href = URL.createObjectURL(blob)
      document.body.appendChild(linkElem)
      linkElem.click()
      document.body.removeChild(linkElem)
    }
  } else {
    console.error(_vxetable.t('vxe.error.notExp'))
  }
}

function replaceDoubleQuotation (val: string) {
  return val.replace(/^"/, '').replace(/"$/, '')
}

function parseCsv (columns: ColumnConfig[], content: string) {
  const list = content.split('\n')
  const fields: string[] = []
  const rows: any[] = []
  if (list.length) {
    const rList = list.slice(1)
    list[0].split(',').map(replaceDoubleQuotation)
    rList.forEach((r) => {
      if (r) {
        const item: { [key: string]: any } = {}
        r.split(',').forEach((val, colIndex) => {
          if (fields[colIndex]) {
            item[fields[colIndex]] = replaceDoubleQuotation(val)
          }
        })
        rows.push(item)
      }
    })
  }
  return { fields, rows }
}

function checkImportData (columns: ColumnConfig[], fields: string[], rows: any[]) {
  let tableFields: string[] = []
  columns.forEach((column) => {
    let field = column.property
    if (field) {
      tableFields.push(field)
    }
  })
  return tableFields.every((field) => fields.includes(field))
}

function importXLSX (params: InterceptorImportParams) {
  const { columns, options, file } = params
  const $table: any = params.$table
  const { _importResolve } = $table
  const fileReader = new FileReader()
  fileReader.onload = (e: any) => {
    const workbook = XLSX.read(e.target.result, { type: 'binary' })
    const csvData: string = XLSX.utils.sheet_to_csv(workbook.Sheets.Sheet1)
    const { fields, rows } = parseCsv(columns, csvData)
    const status = checkImportData(columns, fields, rows)
    if (status) {
      $table.createData(rows)
        .then((data: any[]) => {
          if (options.mode === 'append') {
            $table.insertAt(data, -1)
          } else {
            $table.reloadData(data)
          }
        })
      if (options.message !== false) {
        _vxetable.modal.message({ message: XEUtils.template(_vxetable.t('vxe.table.impSuccess'), [rows.length]), status: 'success' })
      }
    } else if (options.message !== false) {
      _vxetable.modal.message({ message: _vxetable.t('vxe.error.impFields'), status: 'error' })
    }
    if (_importResolve) {
      _importResolve(status)
      $table._importResolve = null
    }
  }
  fileReader.readAsBinaryString(file)
}

function handleImportEvent (params: InterceptorImportParams) {
  if (params.options.type === 'xlsx') {
    importXLSX(params)
    return false
  }
}

function handleExportEvent (params: InterceptorExportParams) {
  if (params.options.type === 'xlsx') {
    exportXLSX(params)
    return false
  }
}

/**
 * 基于 vxe-table 表格的增强插件，支持导出 xlsx 格式
 */
export const VXETablePluginExportXLSX = {
  install (xtable: typeof VXETable) {
    const { interceptor } = xtable
    _vxetable = xtable
    Object.assign(xtable.types, { xlsx: 1 })
    interceptor.mixin({
      'event.import': handleImportEvent,
      'event.export': handleExportEvent
    })
  }
}

if (typeof window !== 'undefined' && window.VXETable) {
  window.VXETable.use(VXETablePluginExportXLSX)
}

export default VXETablePluginExportXLSX
