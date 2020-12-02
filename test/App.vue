<template>
  <div id="app">
    <el-upload
      v-show="!file"
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
    </el-upload>

    <excel-view
      v-if="file || excelData"
      :file="file"
      :excelData="excelData"
      @on-reach-top="reachTop"
    />
  </div>
</template>

<script>
import ExcelView from "../src/components/ExcelView/index";
import axios from "axios";
axios.defaults.withCredentials = true;
export default {
  name: "App",
  components: {
    ExcelView,
  },
  created() {
    // axios
    //   .request({
    //     url: "http://127.0.0.1:8090/excel/part?fileId=6",
    //   })
    //   .then((res) => {
    //     console.info("res", res);
    //   })
    //   .catch((error) => {
    //     console.error("error", error);
    //   });

    axios
      .get("http://127.0.0.1:8090/excel/part?fileId=6", {
        responseType: "blob",
      })
      .then((res) => {
        // console.info("res", res);
        let blob = res.data;
        let reader = new FileReader();
        var that = this;

        reader.onload = (e) => {
          that.excelData = e.target.result;
        };

        reader.readAsBinaryString(blob);
      })
      .catch((error) => {
        console.error("error", error);
      });
  },
  data() {
    return {
      file: null,
      excelData: null,
    };
  },
  methods: {
    reachTop() {
      // alert("reach top");
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
