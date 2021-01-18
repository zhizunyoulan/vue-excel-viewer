<template>
  <div id="app">
    <input @change="chooseFile" type="file"/>

    <excel-viewer
      ref="excelViewer"
      :height="500"
      :first-row-num="firstRowNum"
      :min-col-counts="5"
      :border-collapse="false"
      @on-reach-top="reachTop"
      @on-reach-bottom="reachBottom"
      @on-row-select="onRowSelect"
      @on-col-select="onColSelect"
      @on-cell-select="onCellSelect"
      @on-before-open="beforeOpen"
      @on-after-open="afterOpen"
    />

    <!-- <excel-viewer
      ref="excelSecondViewer"
      :height="500"
      :first-row-num="firstRowNum"
      :min-col-counts="5"
      :border-collapse="false"
      @on-reach-top="reachTop"
      @on-reach-bottom="reachBottom"
      @on-row-select="onRowSelect"
      @on-col-select="onColSelect"
      @on-cell-select="onCellSelect"
      @on-before-open="beforeOpen"
      @on-after-open="afterOpen"
    /> -->
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
      this.$refs.excelViewer.setRowBackgroundColor(2,'red');
      this.$refs.excelViewer.freezeCellAt(3, 1);
    },
    onRowSelect(rowNum, selectRowValues) {//行选择的事件 on row select
      console.info("row select", rowNum, selectRowValues);
      //设置行的背景颜色 set row background-color
      // this.$refs.excelViewer.setSelectedBackgroundColor('red');
    },
    onColSelect(colNum) {//列选择事件 on column select
      console.info("col select", colNum);
    },
    onCellSelect(rowNum, colNum, value) {//单元格选择的事件 on cell select
      //设置背景颜色 set backgroundColor 
      this.$refs.excelViewer.setCellBackgroundColor(rowNum, colNum, 'red');

      var rowValues = this.$refs.excelViewer.getRowValues(rowNum);
      console.info('rowValues', rowValues)

      var cellValue = this.$refs.excelViewer.getCellValue(rowNum, colNum);
      console.info('cellValue', cellValue)

      //冻结窗格 freeze at cell
      // 
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
