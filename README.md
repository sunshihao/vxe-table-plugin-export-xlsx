
# vxe-table-plugin-export-xlsx-xhx

基于 [vxe-table](https://github.com/xuliangzhan/vxe-table) 表格的增强插件，支持导出 xlsx 格式

## introductions

基于 vxe-table-plugin-export-xlsx 1.3.4版本进行工程适应性修改
注意: 隶属于旧项目使用, 新工程请参照 vxe-table-plugin-export-xlsx

## history

版本 1.0.9 对IE11进行适配
版本 1.0.2 不进行多余导出

## Installing
```shell
npm install xe-utils vxe-table vxe-table-plugin-export-xlsx-xhx xlsx
```
```javascript
import Vue from 'vue'
import VXETable from 'vxe-table'
import VXETablePluginExportXLSX from 'vxe-table-plugin-export-xlsx-xhx'
import 'vxe-table/lib/index.css'
Vue.use(VXETable)
VXETable.use(VXETablePluginExportXLSX)
```
## Demo
```html
<vxe-toolbar>
  <template v-slot:buttons>
    <vxe-button @click="exportEvent">导出.xlsx</vxe-button>
  </template>
</vxe-toolbar>
<vxe-table
  border
  ref="xTable"
  height="600"
  :data="tableData">
  <vxe-table-column type="index" width="60"></vxe-table-column>
  <vxe-table-column field="name" title="Name"></vxe-table-column>
  <vxe-table-column field="age" title="Age"></vxe-table-column>
  <vxe-table-column field="date" title="Date"></vxe-table-column>
</vxe-table>
```
```javascript
export default {
  data () {
    return {
      tableData: [
        {
          id: 100,
          name: 'test',
          age: 26,
          date: null
        }
      ]
    }
  },
  methods: {
    exportEvent() {
      this.$refs.xTable.exportData({
        filename: 'export',
        sheetName: 'Sheet1',
        type: 'xlsx'
      })
    }
  }
}
```
## License
MIT License, 2019-present, Xu Liangzhan

