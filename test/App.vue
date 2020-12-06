<template>
  <div id="app">
    <input @change="chooseFile" type="file"/>

    <excel-viewer
      ref="excelViewer"
      :height="300"
      :first-row-index="firstRowIndex"
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
  </div>
</template>

<script>

export default {
  name: "App",
  data() {
    return {
      firstRowIndex: 2
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
    onRowSelect(index, selectRowValues) {//行选择的事件 on row select
      console.info("row select", index, selectRowValues);
      //设置行的背景颜色 set row background-color
      this.$refs.excelViewer.setSelectedBackgroundColor('red');
    },
    onColSelect(index) {//列选择事件 on column select
      console.info("col select", index);
    },
    onCellSelect(rowIndex, colIndex, value) {//单元格选择的事件 on cell select
      //设置背景颜色 set backgroundColor 
      this.$refs.excelViewer.setCellBackgroundColor(rowIndex, colIndex, 'red');
      //冻结窗格 freeze at cell
      this.$refs.excelViewer.freezeCellAt(rowIndex, colIndex);
      if (value) {
        console.info("cell select", rowIndex, colIndex, value);
      } else {
        console.info("cell select， value empty", rowIndex, colIndex);
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
