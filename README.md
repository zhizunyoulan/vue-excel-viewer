## Install
```shell
npm i @uublue/vue-excel-viewer
```
## Props

| name  |  type |  default | description |
| ------------ | ------------ | ------------ | ------------ | 
| height  | Number  | 500  | the height of excel viewer |
|  firstRowIndex |  Number | 1  | 第一行数据的行序号 |
| minColCounts  | Number  |  - | 列的最小显示个数 |


## Example
### 基本用法
```javascript
import VueExcelViewer from '@uublue/vue-excel-viewer'
import '@uublue/vue-excel-viewer/lib/vue-excel-viewer.css'
Vue.use(VueExcelViewer)

```

```html
<template>
  <div id="app">
    <input @change="chooseFile" type="file"/>

    <excel-viewer
      ref="excelViewer"
      :height="300"
      :first-row-num="firstRowNum"
      :min-col-counts="5"
      @on-reach-top="reachTop"
      @on-reach-bottom="reachBottom"
      @on-row-select="onRowSelect"
      @on-col-select="onColSelect"
      @on-cell-select="onCellSelect"
      @on-before-open="beforeOpen"
      @on-after-open="afterOpen"
    />
  </div>
</template>

<script>
export default {
  name: "App",
  data() {
    return {
      firstRowNum: 2
    };
  },
  methods: {
    chooseFile(e){
      console.info("excel file select", e);
      //open excel file
      this.$refs.excelViewer.openExcelFile(e.target.files[0]);
    },
    beforeOpen() {//文件打开前的事件 on before open
      console.info("excel before open");
      
    },
    afterOpen() {//文件打开后的事件 on after open
      console.info("excel after open");
      this.$refs.excelViewer.setRowBackgroundColor(5,'red');
    },
    onRowSelect(rowNum, selectRowValues) {//行选择的事件 on row select
      console.info("row select", rowNum, selectRowValues);
      //设置行的背景颜色 set row background-color
      this.$refs.excelViewer.setSelectedBackgroundColor('red');
    },
    onColSelect(colNum) {//列选择事件 on column select
      console.info("col select", colNum);
    },
    onCellSelect(rowNum, colNum, value) {//单元格选择的事件 on cell select
      //设置背景颜色 set backgroundColor 
      this.$refs.excelViewer.setCellBackgroundColor(rowNum, colNum, 'red');
      //冻结窗格 freeze at cell
      this.$refs.excelViewer.freezeCellAt(rowNum, colNum);
      if (value) {
        console.info("cell select", rowNum, colNum, value);
      } else {
        console.info("cell select， value empty", rowNum, colNum);
      }
    },
    reachTop() {//滚动到顶部的事件 on reach top
      console.info("touch top");
    },
    reachBottom() {//滚动到底部的事件 on reach bottom
      console.info("touch bottom");
    }
  },
};
</script>
```

### 从文件流获取excel
```javascript
        axios
          .get(`/file`, {
            responseType: "blob",
          })
          .then((res) => {
            let blob = res.data;
            let reader = new FileReader();
            var self = this
            reader.onload = (e) => {
              self.$refs.excelView.openExcelData(e.target.result)
            };
            reader.readAsBinaryString(blob);
          })
          .catch((error) => {
            console.error("error", error);
          });
```