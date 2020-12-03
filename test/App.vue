<template>
  <div id="app">
    <!-- <el-upload
      class="upload-demo"
      drag
      action="http://127.0.0.1:8090/upload/excel"
      :with-credentials="true"
      :on-success="onUploadSuccess"
      :http-request="httpRequest"
    >
      <i class="el-icon-upload"></i>
      <div class="el-upload__text">将文件拖到此处，或<em>点击上传</em></div>
      <div class="el-upload__tip" slot="tip">
        只能上传jpg/png文件，且不超过500kb
      </div>
    </el-upload> -->

    <excel-view
      ref="excelView"
      :height="700"
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
// import ExcelView from "../src/components/ExcelView/index";
import axios from "axios";
axios.defaults.withCredentials = true;
export default {
  name: "App",
  components: {
    // ExcelView,
  },
  created() {
    axios
      .get("http://127.0.0.1:8090/excel/part?fileId=8&start=0&end=20", {
        responseType: "blob",
      })
      .then((res) => {
        console.info('res',res)
        let blob = res.data;
        let reader = new FileReader();
        var that = this;

        reader.onload = (e) => {
          that.$refs.excelView.openExcelData(e.target.result)
        };

        reader.readAsBinaryString(blob);
      })
      .catch((error) => {
        console.error("error", error);
      });
  },
  data() {
    return {

    };
  },
  methods: {
    beforeOpen(){
      console.info("excel 准备打开");
    },
    afterOpen(){
      console.info("excel 打开完毕");
      // alert('打开完毕')
    },
    onRowSelect(index){
      console.info("row select", index);
    },
    onColSelect(index){
      console.info("col select", index);
    },
    onCellSelect(rowIndex, colIndex, value){
      console.info("cell select", rowIndex, colIndex, value);
      if(value) {
        console.info("cell value", value);
      }else{
        console.info("cell value empty");
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
      this.file = param.file;
    },
    onUploadSuccess(response, file, fileList) {
      console.info("res", response, file, fileList);
    },
  },
};
</script>
