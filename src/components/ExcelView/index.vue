<template>
  <ul
    class="infinite-list"
    ref="excelPanel"
    v-infinite-scroll="reachBottom"
    style="overflow: auto"
    :style="{ height: height + 'px' }"
  >
    <table class="excel-table" cellspacing="0">
      <thead></thead>
      <tbody></tbody>
    </table>
  </ul>
</template>

<script>
import XLSX from "xlsx";

function transformMerges(merges) {
  var mergesObj = {};
  if (merges && Array.isArray(merges)) {
    merges.forEach((merge) => {
      var mergeItem = { start: merge.s.r + "-" + merge.s.c, invalids: {} };
      var rowspan = merge.e.r - merge.s.r;
      var colspan = merge.e.c - merge.s.c;
      if (rowspan) {
        mergeItem.rowspan = rowspan + 1;
      }
      if (colspan) {
        mergeItem.colspan = colspan + 1;
      }
      for (var i = merge.s.r; i <= merge.e.r; i++) {
        for (var j = merge.s.c; j <= merge.e.c; j++) {
          if (i != merge.s.r || j != merge.s.c) {
            mergeItem.invalids[i + "-" + j] = true;
          }
        }
      }
      mergesObj[mergeItem.start] = mergeItem;
    });
  }

  return mergesObj;
}

function renderSheetAt(domElement, maxRowRange, dataArray, merges) {
  var tableHeadElement = domElement.querySelector("thead");
  var emptyTh = document.createElement("th");
  tableHeadElement.appendChild(emptyTh);
  for (var i = 0; i < maxRowRange; i++) {
    var sign = String.fromCharCode(i + 65);
    var th = document.createElement("th");
    th.innerText = sign;
    tableHeadElement.appendChild(th);
  }
  var tableBodyElement = domElement.querySelector("tbody");
  var invalidCells = {};
  for (var ri = 0; ri < dataArray.length; ri++) {
    var row = dataArray[ri];
    var tr = document.createElement("tr");
    tr.classList.add("excel-row");
    var emptyTd = document.createElement("td");
    emptyTd.innerText = ri + 1;
    tr.appendChild(emptyTd);
    for (var ci = 0; ci < maxRowRange; ci++) {
      var cellKey = ri + "-" + ci;
      // debugger
      // console.info(cellKey, invalidCells[cellKey]);
      if (!invalidCells[cellKey]) {
        var td = document.createElement("td");
        td.classList.add("excel-cell");
        var cellValue = row[ci];
        td.innerText = cellValue || "";
        var merge = merges[cellKey];
        // console.info("merge", cellKey, merges);
        if (merge) {
          invalidCells = Object.assign(invalidCells, merge.invalids);
          if (merge.rowspan) {
            td.setAttribute("rowspan", merge.rowspan);
          }
          if (merge.colspan) {
            td.setAttribute("colspan", merge.colspan);
          }
        }
        tr.appendChild(td);
      }
    }
    tableBodyElement.appendChild(tr);
  }
}

function renderWorkbookSheet(workbook, sheetName) {
  var worksheet = workbook.Sheets[sheetName];
  console.info("renderWorkbookSheet", worksheet);
  var defaultRange = worksheet["!ref"];
  var lastCellPosition = defaultRange.split(":")[1];
  var lastCellPositionMatchInfo = lastCellPosition.match(/(\D+)(\d+)/);
  console.info("lastCellPositionMatchInfo", lastCellPositionMatchInfo);
  var range = "A1:" + lastCellPositionMatchInfo[1] + 50;
  var merges = worksheet["!merges"];
  // console.info("range", range, merges);
  var sheetDatas = XLSX.utils.sheet_to_json(worksheet, {
    raw: false,
    header: 1,
    range: range,
  });
  console.info("workbook sheet data", sheetDatas);
  var tableElement = document.querySelector(".excel-table");

  var maxRowRange = parseInt(lastCellPositionMatchInfo[2]);
  console.info("maxRowRange", maxRowRange);
  renderSheetAt(tableElement, maxRowRange, sheetDatas, transformMerges(merges));
}

export default {
  data() {
    return {
      isScrollAtTop: true,
    };
  },
  props: {
    excelFile: File,
    excelData: String,
    height: {
      type: Number,
      default: 500,
    },
  },
  created() {},
  mounted() {
    if (this.excelData) {
      this.showExcel();
    } else if (this.excelFile) {
      this.showExcel();
    }

    var that = this;
    this.$refs.excelPanel.onscroll = function (e) {
      var scrollTop = e.target.scrollTop;
      if (scrollTop == 0) {
        if (!this.isScrollAtTop) {
          that.$emit("on-reach-top");
        }
        this.isScrollAtTop = true;
      } else {
        this.isScrollAtTop = false;
      }
    };
  },
  methods: {
    showExcel() {
      if (this.excelData) {
        // console.info("start read workbook", this.excelData);
        var workbook = XLSX.read(this.excelData, {
          type: "binary",
        });
        // console.info('workbook',workbook)
        var sheetName = workbook.SheetNames[0];
        if (sheetName) {
          renderWorkbookSheet(workbook, sheetName);
        }
      } else if (this.excelFile) {
        var that = this;
        const reader = new FileReader();
        reader.onload = function (e) {
          const data = e.target.result;
          console.info("start read workbook");
          var workbook = XLSX.read(data, {
            type: "binary",
          });
          // console.info('workbook',workbook)
          var sheetName = workbook.SheetNames[0];
          if (sheetName) {
            renderWorkbookSheet(workbook, sheetName);
          }
        };
        reader.readAsBinaryString(this.excelFile);
      }
    },
    reachBottom() {
      this.$emit("on-reach-bottom");
      console.info("touch bottom");
    },
  },
};
</script>

<style>
.excel-table:hover {
  cursor: pointer;
}
.excel table tr td {
  border: 1px solid black;
}
.excel-row {
  height: 30px;
}
.excel-cell {
  border: 1px solid black;
  width: 100px;
}
</style>