function _typeof(obj) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) { return typeof obj; } : function (obj) { return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }, _typeof(obj); }
(function (global, factory) {
  if (typeof define === "function" && define.amd) {
    define("vxe-table-plugin-export-xlsx", ["exports", "xe-utils", "exceljs"], factory);
  } else if (typeof exports !== "undefined") {
    factory(exports, require("xe-utils"), require("exceljs"));
  } else {
    var mod = {
      exports: {}
    };
    factory(mod.exports, global.XEUtils, global.ExcelJS);
    global.VXETablePluginExportXLSX = mod.exports.default;
  }
})(typeof globalThis !== "undefined" ? globalThis : typeof self !== "undefined" ? self : this, function (_exports, _xeUtils, ExcelJS) {
  "use strict";

  Object.defineProperty(_exports, "__esModule", {
    value: true
  });
  _exports["default"] = _exports.VXETablePluginExportXLSX = void 0;
  _xeUtils = _interopRequireDefault(_xeUtils);
  ExcelJS = _interopRequireWildcard(ExcelJS);
  function _getRequireWildcardCache(nodeInterop) { if (typeof WeakMap !== "function") return null; var cacheBabelInterop = new WeakMap(); var cacheNodeInterop = new WeakMap(); return (_getRequireWildcardCache = function _getRequireWildcardCache(nodeInterop) { return nodeInterop ? cacheNodeInterop : cacheBabelInterop; })(nodeInterop); }
  function _interopRequireWildcard(obj, nodeInterop) { if (!nodeInterop && obj && obj.__esModule) { return obj; } if (obj === null || _typeof(obj) !== "object" && typeof obj !== "function") { return { "default": obj }; } var cache = _getRequireWildcardCache(nodeInterop); if (cache && cache.has(obj)) { return cache.get(obj); } var newObj = {}; var hasPropertyDescriptor = Object.defineProperty && Object.getOwnPropertyDescriptor; for (var key in obj) { if (key !== "default" && Object.prototype.hasOwnProperty.call(obj, key)) { var desc = hasPropertyDescriptor ? Object.getOwnPropertyDescriptor(obj, key) : null; if (desc && (desc.get || desc.set)) { Object.defineProperty(newObj, key, desc); } else { newObj[key] = obj[key]; } } } newObj["default"] = obj; if (cache) { cache.set(obj, newObj); } return newObj; }
  function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { "default": obj }; }
  var defaultHeaderBackgroundColor = 'f8f8f9';
  var defaultCellFontColor = '606266';
  var defaultCellBorderStyle = 'thin';
  var defaultCellBorderColor = 'e8eaec';
  function getCellLabel(column, cellValue) {
    if (cellValue) {
      switch (column.cellType) {
        case 'string':
          return _xeUtils["default"].toValueString(cellValue);
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
  function getFooterData(opts, footerData) {
    var footerFilterMethod = opts.footerFilterMethod;
    return footerFilterMethod ? footerData.filter(function (items, index) {
      return footerFilterMethod({
        items: items,
        $rowIndex: index
      });
    }) : footerData;
  }
  function getFooterCellValue($table, opts, rows, column) {
    var cellValue = getCellLabel(column, rows[$table.getVMColumnIndex(column)]);
    return cellValue;
  }
  function getValidColumn(column) {
    var childNodes = column.childNodes;
    var isColGroup = childNodes && childNodes.length;
    if (isColGroup) {
      return getValidColumn(childNodes[0]);
    }
    return column;
  }
  function setExcelRowHeight(excelRow, height) {
    if (height) {
      excelRow.height = _xeUtils["default"].floor(height * 0.75, 12);
    }
  }
  function setExcelCellStyle(excelCell, align) {
    excelCell.protection = {
      locked: false
    };
    excelCell.alignment = {
      vertical: 'middle',
      horizontal: align || 'left'
    };
  }
  function getDefaultBorderStyle() {
    return {
      top: {
        style: defaultCellBorderStyle,
        color: {
          argb: defaultCellBorderColor
        }
      },
      left: {
        style: defaultCellBorderStyle,
        color: {
          argb: defaultCellBorderColor
        }
      },
      bottom: {
        style: defaultCellBorderStyle,
        color: {
          argb: defaultCellBorderColor
        }
      },
      right: {
        style: defaultCellBorderStyle,
        color: {
          argb: defaultCellBorderColor
        }
      }
    };
  }
  function exportXLSX(params) {
    var msgKey = 'xlsx';
    var $table = params.$table,
      options = params.options,
      columns = params.columns,
      colgroups = params.colgroups,
      datas = params.datas;
    var $vxe = $table.$vxe,
      rowHeight = $table.rowHeight,
      allHeaderAlign = $table.headerAlign,
      allAlign = $table.align,
      allFooterAlign = $table.footerAlign;
    var modal = $vxe.modal,
      t = $vxe.t;
    var message = options.message,
      sheetName = options.sheetName,
      isHeader = options.isHeader,
      isFooter = options.isFooter,
      isMerge = options.isMerge,
      isColgroup = options.isColgroup,
      original = options.original,
      useStyle = options.useStyle,
      sheetMethod = options.sheetMethod;
    var showMsg = message !== false;
    var mergeCells = $table.getMergeCells();
    var colList = [];
    var footList = [];
    var sheetCols = [];
    var sheetMerges = [];
    var beforeRowCount = 0;
    var colHead = {};
    columns.forEach(function (column) {
      var id = column.id,
        property = column.property,
        renderWidth = column.renderWidth;
      colHead[id] = original ? property : column.getTitle();
      sheetCols.push({
        key: id,
        width: _xeUtils["default"].ceil(renderWidth / 8, 1)
      });
    });
    // 处理表头
    if (isHeader) {
      // 处理分组
      if (isColgroup && !original && colgroups) {
        colgroups.forEach(function (cols, rIndex) {
          var groupHead = {};
          columns.forEach(function (column) {
            groupHead[column.id] = null;
          });
          cols.forEach(function (column) {
            var _colSpan = column._colSpan,
              _rowSpan = column._rowSpan;
            var validColumn = getValidColumn(column);
            var columnIndex = columns.indexOf(validColumn);
            groupHead[validColumn.id] = original ? validColumn.property : column.getTitle();
            if (_colSpan > 1 || _rowSpan > 1) {
              sheetMerges.push({
                s: {
                  r: rIndex,
                  c: columnIndex
                },
                e: {
                  r: rIndex + _rowSpan - 1,
                  c: columnIndex + _colSpan - 1
                }
              });
            }
          });
          colList.push(groupHead);
        });
      } else {
        colList.push(colHead);
      }
      beforeRowCount += colList.length;
    }
    // 处理合并
    if (isMerge && !original) {
      mergeCells.forEach(function (mergeItem) {
        var mergeRowIndex = mergeItem.row,
          mergeRowspan = mergeItem.rowspan,
          mergeColIndex = mergeItem.col,
          mergeColspan = mergeItem.colspan;
        sheetMerges.push({
          s: {
            r: mergeRowIndex + beforeRowCount,
            c: mergeColIndex
          },
          e: {
            r: mergeRowIndex + beforeRowCount + mergeRowspan - 1,
            c: mergeColIndex + mergeColspan - 1
          }
        });
      });
    }
    var rowList = datas.map(function (item) {
      var rest = {};
      columns.forEach(function (column) {
        rest[column.id] = getCellLabel(column, item[column.id]);
      });
      return rest;
    });
    beforeRowCount += rowList.length;
    // 处理表尾
    if (isFooter) {
      var _$table$getTableData = $table.getTableData(),
        footerData = _$table$getTableData.footerData;
      var footers = getFooterData(options, footerData);
      var mergeFooterItems = $table.getMergeFooterItems();
      // 处理合并
      if (isMerge && !original) {
        mergeFooterItems.forEach(function (mergeItem) {
          var mergeRowIndex = mergeItem.row,
            mergeRowspan = mergeItem.rowspan,
            mergeColIndex = mergeItem.col,
            mergeColspan = mergeItem.colspan;
          sheetMerges.push({
            s: {
              r: mergeRowIndex + beforeRowCount,
              c: mergeColIndex
            },
            e: {
              r: mergeRowIndex + beforeRowCount + mergeRowspan - 1,
              c: mergeColIndex + mergeColspan - 1
            }
          });
        });
      }
      footers.forEach(function (rows) {
        var item = {};
        columns.forEach(function (column) {
          item[column.id] = getFooterCellValue($table, options, rows, column);
        });
        footList.push(item);
      });
    }
    var exportMethod = function exportMethod() {
      var workbook = new ExcelJS.Workbook();
      var sheet = workbook.addWorksheet(sheetName);
      workbook.creator = 'vxe-table';
      sheet.columns = sheetCols;
      if (isHeader) {
        sheet.addRows(colList).forEach(function (excelRow) {
          if (useStyle) {
            setExcelRowHeight(excelRow, rowHeight);
          }
          excelRow.eachCell(function (excelCell) {
            var excelCol = sheet.getColumn(excelCell.col);
            var column = $table.getColumnById(excelCol.key);
            var headerAlign = column.headerAlign,
              align = column.align;
            setExcelCellStyle(excelCell, headerAlign || align || allHeaderAlign || allAlign);
            if (useStyle) {
              Object.assign(excelCell, {
                font: {
                  bold: true,
                  color: {
                    argb: defaultCellFontColor
                  }
                },
                fill: {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: {
                    argb: defaultHeaderBackgroundColor
                  }
                },
                border: getDefaultBorderStyle()
              });
            }
          });
        });
      }
      sheet.addRows(rowList).forEach(function (excelRow) {
        if (useStyle) {
          setExcelRowHeight(excelRow, rowHeight);
        }
        excelRow.eachCell(function (excelCell) {
          var excelCol = sheet.getColumn(excelCell.col);
          var column = $table.getColumnById(excelCol.key);
          var align = column.align;
          setExcelCellStyle(excelCell, align || allAlign);
          if (useStyle) {
            Object.assign(excelCell, {
              font: {
                color: {
                  argb: defaultCellFontColor
                }
              },
              border: getDefaultBorderStyle()
            });
          }
        });
      });
      if (isFooter) {
        sheet.addRows(footList).forEach(function (excelRow) {
          if (useStyle) {
            setExcelRowHeight(excelRow, rowHeight);
          }
          excelRow.eachCell(function (excelCell) {
            var excelCol = sheet.getColumn(excelCell.col);
            var column = $table.getColumnById(excelCol.key);
            var footerAlign = column.footerAlign,
              align = column.align;
            setExcelCellStyle(excelCell, footerAlign || align || allFooterAlign || allAlign);
            if (useStyle) {
              Object.assign(excelCell, {
                font: {
                  color: {
                    argb: defaultCellFontColor
                  }
                },
                border: getDefaultBorderStyle()
              });
            }
          });
        });
      }
      if (useStyle && sheetMethod) {
        var sParams = {
          options: options,
          workbook: workbook,
          worksheet: sheet,
          columns: columns,
          colgroups: colgroups,
          datas: datas,
          $table: $table
        };
        sheetMethod(sParams);
      }
      sheetMerges.forEach(function (_ref) {
        var s = _ref.s,
          e = _ref.e;
        sheet.mergeCells(s.r + 1, s.c + 1, e.r + 1, e.c + 1);
      });
      workbook.xlsx.writeBuffer().then(function (buffer) {
        var blob = new Blob([buffer], {
          type: 'application/octet-stream'
        });
        // 导出 xlsx
        downloadFile(params, blob, options);
        if (showMsg && modal) {
          modal.close(msgKey);
          modal.message({
            content: t('vxe.table.expSuccess'),
            status: 'success'
          });
        }
      });
    };
    if (showMsg && modal) {
      modal.message({
        id: msgKey,
        content: t('vxe.table.expLoading'),
        status: 'loading',
        duration: -1
      });
      setTimeout(exportMethod, 1500);
    } else {
      exportMethod();
    }
  }
  function downloadFile(params, blob, options) {
    var $table = params.$table;
    var $vxe = $table.$vxe;
    var modal = $vxe.modal,
      t = $vxe.t;
    var message = options.message,
      filename = options.filename,
      type = options.type;
    var showMsg = message !== false;
    if (window.Blob) {
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
      if (showMsg && modal) {
        modal.alert({
          content: t('vxe.error.notExp'),
          status: 'error'
        });
      }
    }
  }
  function checkImportData(tableFields, fields) {
    return fields.some(function (field) {
      return tableFields.indexOf(field) > -1;
    });
  }
  function importError(params) {
    var $table = params.$table,
      options = params.options;
    var $vxe = $table.$vxe,
      _importReject = $table._importReject;
    var showMsg = options.message !== false;
    var modal = $vxe.modal,
      t = $vxe.t;
    if (showMsg && modal) {
      modal.message({
        content: t('vxe.error.impFields'),
        status: 'error'
      });
    }
    if (_importReject) {
      _importReject({
        status: false
      });
    }
  }
  function importXLSX(params) {
    var $table = params.$table,
      columns = params.columns,
      options = params.options,
      file = params.file;
    var $vxe = $table.$vxe,
      _importResolve = $table._importResolve;
    var modal = $vxe.modal,
      t = $vxe.t;
    var showMsg = options.message !== false;
    var fileReader = new FileReader();
    fileReader.onerror = function () {
      importError(params);
    };
    fileReader.onload = function (evnt) {
      var tableFields = [];
      columns.forEach(function (column) {
        var field = column.property;
        if (field) {
          tableFields.push(field);
        }
      });
      var workbook = new ExcelJS.Workbook();
      var readerTarget = evnt.target;
      if (readerTarget) {
        workbook.xlsx.load(readerTarget.result).then(function (wb) {
          var firstSheet = wb.worksheets[0];
          if (firstSheet) {
            var sheetValues = firstSheet.getSheetValues();
            var fieldIndex = _xeUtils["default"].findIndexOf(sheetValues, function (list) {
              return list && list.length > 0;
            });
            var fields = sheetValues[fieldIndex];
            var status = checkImportData(tableFields, fields);
            if (status) {
              var records = sheetValues.slice(fieldIndex).map(function (list) {
                var item = {};
                list.forEach(function (cellValue, cIndex) {
                  item[fields[cIndex]] = cellValue;
                });
                var record = {};
                tableFields.forEach(function (field) {
                  record[field] = _xeUtils["default"].isUndefined(item[field]) ? null : item[field];
                });
                return record;
              });
              $table.createData(records).then(function (data) {
                var loadRest;
                if (options.mode === 'insert') {
                  loadRest = $table.insertAt(data, -1);
                } else {
                  loadRest = $table.reloadData(data);
                }
                return loadRest.then(function () {
                  if (_importResolve) {
                    _importResolve({
                      status: true
                    });
                  }
                });
              });
              if (showMsg && modal) {
                modal.message({
                  content: t('vxe.table.impSuccess', [records.length]),
                  status: 'success'
                });
              }
            } else {
              importError(params);
            }
          } else {
            importError(params);
          }
        });
      } else {
        importError(params);
      }
    };
    fileReader.readAsArrayBuffer(file);
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
    install: function install(vxetable) {
      var interceptor = vxetable.interceptor;
      vxetable.setup({
        "export": {
          types: {
            xlsx: 0
          }
        }
      });
      interceptor.mixin({
        'event.import': handleImportEvent,
        'event.export': handleExportEvent
      });
    }
  };
  _exports.VXETablePluginExportXLSX = VXETablePluginExportXLSX;
  if (typeof window !== 'undefined' && window.VXETable && window.VXETable.use) {
    window.VXETable.use(VXETablePluginExportXLSX);
  }
  var _default = VXETablePluginExportXLSX;
  _exports["default"] = _default;
});