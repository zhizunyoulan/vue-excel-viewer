<template>
  <ul
    class="infinite-list"
    ref="excelPanel"
    v-infinite-scroll="reachBottom"
    style="overflow: auto"
    :style="{ height: height + 'px' }"
  >
    <table class="excel-table">
      <tbody></tbody>
    </table>
  </ul>
</template>

<script>
import XLSX from "xlsx";

function transformMerges(merges) {
  var mergesObj ={};
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
    mergesObj[mergeItem.start] = mergeItem
  });

  return mergesObj;
}

function renderSheetAt(domElement, dataArray, merges) {
  var invalidCells = {};
  for (var ri = 0; ri < dataArray.length; ri++) {
    var row = dataArray[ri];
    var tr = document.createElement("tr");
    tr.classList.add('excel-row')
    for (var ci = 0; ci < row.length; ci++) {
      var cellKey = ri + "-" + ci;
      // debugger
      if (!invalidCells[cellKey]) {
        console.info(cellKey,invalidCells[cellKey])
        var td = document.createElement("td");
        td.classList.add('excel-cell')
        var cellValue = row[ci];
        td.innerText = cellValue;
        var merge = merges[cellKey];
        console.info('merge',cellKey, merges)
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
    domElement.appendChild(tr);
  }
}

function renderWorkbookSheet(workbook, sheetName) {
  console.info('start renderWorkbookSheet')
  var worksheet = workbook.Sheets[sheetName];
  var defaultRange = worksheet['!ref']
  var lastCellPosition = defaultRange.split(':')[1]
  var lastCharPosition = lastCellPosition.slice(0,1)
  var range = 'A1:' + lastCharPosition + 30
  var merges = worksheet["!merges"];
  console.info('range',range,merges)
  var sheetDatas = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: range });




  var tableBodyElement = document.querySelector(".excel-table tbody");

  renderSheetAt(tableBodyElement, sheetDatas, transformMerges(merges));
}

export default {
  data() {
    return {
      excelFile: null,
      isScrollAtTop: true,
    };
  },
  props: {
    file: File,
    height: {
      type: Number,
      default: 500,
    },
  },
  created() {
    if (this.file) {
      this.excelFile = this.file;
      this.showExcel();
    }
  },
  mounted() {
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
      if (this.excelFile) {
        var that = this;
        const reader = new FileReader();
        reader.onload = function (e) {
          const data = e.target.result;
          console.info('start read workbook')
          var workbook = XLSX.read(data, {
            type: "binary",
          });
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
.excel table tr td {
  border: 1px solid black;
}
.excel-cell {
  border: 1px solid black;
}
</style>