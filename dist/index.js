function _typeof(obj) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) { return typeof obj; } : function (obj) { return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }, _typeof(obj); }
(function (global, factory) {
  if (typeof define === "function" && define.amd) {
    define("vxe-table-plugin-export-xlsx-xhx", ["exports", "xe-utils", "xlsx-js-style"], factory);
  } else if (typeof exports !== "undefined") {
    factory(exports, require("xe-utils"), require("xlsx-js-style"));
  } else {
    var mod = {
      exports: {}
    };
    factory(mod.exports, global.XEUtils, global.xlsxJsStyle);
    global.VXETablePluginExportXLSX = mod.exports.default;
  }
})(typeof globalThis !== "undefined" ? globalThis : typeof self !== "undefined" ? self : this, function (_exports, _xeUtils, XLSX) {
  "use strict";

  Object.defineProperty(_exports, "__esModule", {
    value: true
  });
  _exports["default"] = _exports.VXETablePluginExportXLSX = void 0;
  _xeUtils = _interopRequireDefault(_xeUtils);
  XLSX = _interopRequireWildcard(XLSX);
  function _getRequireWildcardCache(nodeInterop) { if (typeof WeakMap !== "function") return null; var cacheBabelInterop = new WeakMap(); var cacheNodeInterop = new WeakMap(); return (_getRequireWildcardCache = function _getRequireWildcardCache(nodeInterop) { return nodeInterop ? cacheNodeInterop : cacheBabelInterop; })(nodeInterop); }
  function _interopRequireWildcard(obj, nodeInterop) { if (!nodeInterop && obj && obj.__esModule) { return obj; } if (obj === null || _typeof(obj) !== "object" && typeof obj !== "function") { return { "default": obj }; } var cache = _getRequireWildcardCache(nodeInterop); if (cache && cache.has(obj)) { return cache.get(obj); } var newObj = {}; var hasPropertyDescriptor = Object.defineProperty && Object.getOwnPropertyDescriptor; for (var key in obj) { if (key !== "default" && Object.prototype.hasOwnProperty.call(obj, key)) { var desc = hasPropertyDescriptor ? Object.getOwnPropertyDescriptor(obj, key) : null; if (desc && (desc.get || desc.set)) { Object.defineProperty(newObj, key, desc); } else { newObj[key] = obj[key]; } } } newObj["default"] = obj; if (cache) { cache.set(obj, newObj); } return newObj; }
  function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }
  /* eslint-disable no-unused-vars */

  // import XLSX from 'xlsx'

  /* eslint-enable no-unused-vars */
  var _vxetable;
  function getCellLabel(column, cellValue) {
    if (cellValue) {
      switch (column.cellType) {
        case 'string':
          return _xeUtils["default"].toString(cellValue);
        case 'number':
          if (!isNaN(cellValue)) {
            return Number(cellValue);
          }
          break;
        default:
          if (cellValue.length < 12 && !isNaN(cellValue)) {
            return Number(cellValue);
          }
          break;
      }
    }
    return cellValue;
  }
  function getFooterCellValue($table, opts, rows, column) {
    var cellValue = _xeUtils["default"].toString(rows[$table.$getColumnIndex(column)]);
    return cellValue;
  }
  function toBuffer(wbout) {
    var buf = new ArrayBuffer(wbout.length);
    var view = new Uint8Array(buf);
    for (var index = 0; index !== wbout.length; ++index) {
      view[index] = wbout.charCodeAt(index) & 0xFF;
    }
    return buf;
  }
  function exportXLSX(params) {
    var $table = params.$table,
      options = params.options,
      columns = params.columns,
      datas = params.datas;
    var sheetName = options.sheetName,
      isHeader = options.isHeader,
      isFooter = options.isFooter,
      original = options.original,
      message = options.message,
      footerFilterMethod = options.footerFilterMethod;
    var colHead = {};
    var footList = [];
    // const rowList = datas
    if (isHeader) {
      columns.forEach(function (column) {
        colHead[column.id] = _xeUtils["default"].toString(original ? column.property : column.getTitle());
      });
    }
    // 新增部分
    var rowList = datas.map(function (item) {
      var rest = {};
      columns.forEach(function (column) {
        rest[column.id] = getCellLabel(column, item[column.id]);
      });
      return rest;
    });
    if (isFooter) {
      var _$table$getTableData = $table.getTableData(),
        footerData = _$table$getTableData.footerData;
      var footers = footerFilterMethod ? footerData.filter(footerFilterMethod) : footerData;
      footers.forEach(function (rows) {
        var item = {};
        columns.forEach(function (column) {
          item[column.id] = getFooterCellValue($table, options, rows, column);
        });
        footList.push(item);
      });
    }
    var book = XLSX.utils.book_new();
    var sheet = XLSX.utils.json_to_sheet((isHeader ? [colHead] : []).concat(rowList).concat(footList), {
      skipHeader: true
    });
    Object.keys(sheet).forEach(function (key) {
      // 非!开头的属性都是单元格
      if (!key.startsWith('!')) {
        if (key.replace(/[^\d]/g, '') === '1') {
          sheet[key].v = sheet[key].v.padEnd(10, ' '); // 手动补位
          sheet[key].s = {
            font: {
              sz: '10',
              bold: true
            },
            fill: {
              fgColor: {
                rgb: 'F2F2F2'
              }
            },
            border: {
              top: {
                style: 'thin'
              },
              right: {
                style: 'thin'
              },
              bottom: {
                style: 'thin'
              },
              left: {
                style: 'thin'
              }
            }
          };
        } else {
          sheet[key].s = {
            font: {
              sz: '10'
            },
            alignment: {
              wrapText: true,
              vertical: 'top'
            },
            border: {
              top: {
                style: 'thin'
              },
              right: {
                style: 'thin'
              },
              bottom: {
                style: 'thin'
              },
              left: {
                style: 'thin'
              }
            }
          };
        }
      }
    });
    // 转换数据
    XLSX.utils.book_append_sheet(book, sheet, sheetName);
    var wbout = XLSX.write(book, {
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary'
    });
    var blob = new Blob([toBuffer(wbout)], {
      type: 'application/octet-stream'
    });
    // 保存导出
    downloadFile(blob, options);
    if (message !== false) {
      _vxetable.modal.message({
        message: _vxetable.t('vxe.table.expSuccess'),
        status: 'success'
      });
    }
  }
  function downloadFile(blob, options) {
    if (window.Blob) {
      var filename = options.filename,
        type = options.type;
      if (navigator.msSaveBlob) {
        navigator.msSaveBlob(blob, "".concat(filename, ".").concat(type));
      } else {
        var linkElem = document.createElement('a');
        linkElem.target = '_blank';
        linkElem.download = "".concat(filename, ".").concat(type);
        linkElem.href = URL.createObjectURL(blob);
        document.body.appendChild(linkElem);
        linkElem.click();
        document.body.removeChild(linkElem);
      }
    } else {
      console.error(_vxetable.t('vxe.error.notExp'));
    }
  }
  function replaceDoubleQuotation(val) {
    return val.replace(/^"/, '').replace(/"$/, '');
  }
  function parseCsv(columns, content) {
    var list = content.split('\n');
    var fields = [];
    var rows = [];
    if (list.length) {
      var rList = list.slice(1);
      list[0].split(',').map(replaceDoubleQuotation);
      rList.forEach(function (r) {
        if (r) {
          var item = {};
          r.split(',').forEach(function (val, colIndex) {
            if (fields[colIndex]) {
              item[fields[colIndex]] = replaceDoubleQuotation(val);
            }
          });
          rows.push(item);
        }
      });
    }
    return {
      fields: fields,
      rows: rows
    };
  }
  function checkImportData(columns, fields, rows) {
    var tableFields = [];
    columns.forEach(function (column) {
      var field = column.property;
      if (field) {
        tableFields.push(field);
      }
    });
    return tableFields.every(function (field) {
      return fields.includes(field);
    });
  }
  function importXLSX(params) {
    var columns = params.columns,
      options = params.options,
      file = params.file;
    var $table = params.$table;
    var _importResolve = $table._importResolve;
    var fileReader = new FileReader();
    fileReader.onload = function (e) {
      var workbook = XLSX.read(e.target.result, {
        type: 'binary'
      });
      var csvData = XLSX.utils.sheet_to_csv(workbook.Sheets.Sheet1);
      var _parseCsv = parseCsv(columns, csvData),
        fields = _parseCsv.fields,
        rows = _parseCsv.rows;
      var status = checkImportData(columns, fields, rows);
      if (status) {
        $table.createData(rows).then(function (data) {
          if (options.mode === 'append') {
            $table.insertAt(data, -1);
          } else {
            $table.reloadData(data);
          }
        });
        if (options.message !== false) {
          _vxetable.modal.message({
            message: _xeUtils["default"].template(_vxetable.t('vxe.table.impSuccess'), [rows.length]),
            status: 'success'
          });
        }
      } else if (options.message !== false) {
        _vxetable.modal.message({
          message: _vxetable.t('vxe.error.impFields'),
          status: 'error'
        });
      }
      if (_importResolve) {
        _importResolve(status);
        $table._importResolve = null;
      }
    };
    fileReader.readAsBinaryString(file);
  }
  function handleImportEvent(params) {
    if (params.options.type === 'xlsx') {
      importXLSX(params);
      return false;
    }
  }
  function handleExportEvent(params) {
    if (params.options.type === 'xlsx') {
      exportXLSX(params);
      return false;
    }
  }
  /**
   * 基于 vxe-table 表格的增强插件，支持导出 xlsx 格式
   */
  var VXETablePluginExportXLSX = {
    install: function install(xtable) {
      var interceptor = xtable.interceptor;
      _vxetable = xtable;
      Object.assign(xtable.types, {
        xlsx: 1
      });
      interceptor.mixin({
        'event.import': handleImportEvent,
        'event.export': handleExportEvent
      });
    }
  };
  _exports.VXETablePluginExportXLSX = VXETablePluginExportXLSX;
  if (typeof window !== 'undefined' && window.VXETable) {
    window.VXETable.use(VXETablePluginExportXLSX);
  }
  var _default = VXETablePluginExportXLSX;
  _exports["default"] = _default;
});