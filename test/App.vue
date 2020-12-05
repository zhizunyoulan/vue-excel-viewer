<template>
  <div id="app">
    <input @change="choose" type="file"/>

    <excel-view
      ref="excelView"
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
    choose(e){
      console.info("excel 准备打开", e);
      console.info("param", e.target.files);
      this.$refs.excelView.openExcelFile(e.target.files[0]);
    },
    beforeOpen() {
      console.info("excel 准备打开");
      
    },
    afterOpen() {
      console.info("excel 打开完毕");
      this.$refs.excelView.setRowBackgroundColor(5,'red');
    },
    onRowSelect(index, selectRowValues) {
      console.info("row select", index, selectRowValues);
      this.$refs.excelView.setSelectedBackgroundColor('red');
    },
    onColSelect(index) {
      console.info("col select", index);
    },
    onCellSelect(rowIndex, colIndex, value) {
      // this.$refs.excelView.setCellBackgroundColor(rowIndex, colIndex, 'red');
      this.$refs.excelView.freezeCellAt(rowIndex, colIndex);

      if (value) {
        console.info("cell select", rowIndex, colIndex, value);
      } else {
        console.info("cell select， value empty", rowIndex, colIndex);
      }
    },
    reachTop() {
      console.info("touch top");
    },
    reachBottom() {
      console.info("touch bottom");
    },
    httpRequest(param) {
      console.info("param", param);
      this.$refs.excelView.openExcelFile(param.file);
    },
    onUploadSuccess(response, file, fileList) {
      console.info("res", response, file, fileList);
    },
  },
};
</script>
