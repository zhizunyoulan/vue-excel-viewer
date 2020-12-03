<template>
  <ul
    class="infinite-list"
    ref="excelPanel"
    v-infinite-scroll="reachBottom"
    style="overflow: auto"
    :style="{ height: height + 'px' }"
  >
    <table
      class="excel-table"
      cellspacing="0"
      cellpadding="0"
      v-on:click="excelClick"
    >
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
    var sign = parseNumToChars(i);
    var th = document.createElement("th");
    th.classList.add("excel-head-th");
    th.classList.add("excel-cell-col-" + (i + 1));
    th.setAttribute("col-index", i + 1);
    th.setAttribute("col-sign", sign);
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
    emptyTd.classList.add("excel-left-num");
    var rowIndex = ri + 1;
    emptyTd.setAttribute("row-index", rowIndex);
    emptyTd.innerText = rowIndex;
    tr.appendChild(emptyTd);
    for (var ci = 0; ci < maxRowRange; ci++) {
      var colIndex = ci + 1;
      var cellKey = ri + "-" + ci;
      // debugger
      // console.info(cellKey, invalidCells[cellKey]);
      if (!invalidCells[cellKey]) {
        var td = document.createElement("td");
        td.classList.add("excel-cell");
        td.setAttribute("row-index", rowIndex);
        td.setAttribute("col-index", colIndex);
        td.classList.add("excel-cell-col-" + colIndex);
        td.classList.add("excel-cell-row-" + rowIndex);
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
  // console.info("renderWorkbookSheet", worksheet);
  var defaultRange = worksheet["!ref"];
  var lastCellPosition = defaultRange.split(":")[1];
  var lastCellPositionMatchInfo = lastCellPosition.match(/(\D+)(\d+)/);
  // console.info("lastCellPositionMatchInfo", lastCellPositionMatchInfo);
  var range = "A1:" + lastCellPositionMatchInfo[1] + 10;
  var merges = worksheet["!merges"];
  // console.info("range", range, merges);
  var sheetDatas = XLSX.utils.sheet_to_json(worksheet, {
    raw: false,
    header: 1,
    range: range,
  });
  // console.info("workbook sheet data", sheetDatas);
  var tableElement = document.querySelector(".excel-table");

  var maxRowRange = parseCharsTo10(lastCellPositionMatchInfo[1]);
  // console.info("maxRowRange", maxRowRange);
  renderSheetAt(tableElement, maxRowRange, sheetDatas, transformMerges(merges));
}

function parseNumToChars(num) {
  var numBy26Array = [];
  function divide26(num) {
    var shang = Math.floor(num / 26);
    var yu = num % 26;
    if (shang == 0) {
      numBy26Array.push(yu);
    } else if (shang > 0) {
      numBy26Array.push(shang - 1);
      divide26(yu);
    }
  }
  if (!isNaN(num) && num >= 0) {
    divide26(num);
    var result = numBy26Array.reduce(
      (total, currentValue, currentIndex, arr) => {
        return total + String.fromCharCode(currentValue + 65);
      },
      ""
    );

    return result;
  } else {
    throw "必须是大于等于0的数字";
  }
}

function parseCharsTo10(value) {
  if (/\D/.test(value)) {
    var charNumArray = [];
    for (var i = 0; i < value.length; i++) {
      charNumArray.push(value.charCodeAt(i) - 64);
    }
    var result = charNumArray
      .reverse()
      .reduce((total, currentValue, currentIndex, arr) => {
        return total + currentValue * Math.pow(26, currentIndex);
      }, 0);
    return result;
  } else {
    throw "不是字母字符串";
  }
}

export default {
  name: 'ExcelView',
  automount: true,
  data() {
    return {
      isScrollAtTop: true,
    };
  },
  props: {
    height: {
      type: Number,
      default: 500,
    },
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
    excelClick(e) {
      var target = e.target;
      if (target.classList.contains("excel-head-th")) {
        document.querySelectorAll(".selected").forEach((ele) => {
          ele.classList.remove("selected");
          ele.classList.remove("selected-col");
          ele.classList.remove("selected-row");
        });
        target.classList.add("selected");
        var colIndex = target.getAttribute("col-index");
        document
          .querySelectorAll(".excel-cell.excel-cell-col-" + colIndex)
          .forEach((ele) => {
            ele.classList.add("selected");
            ele.classList.add("selected-col");
          });
        this.$emit("on-col-select", colIndex);
      } else if (target.classList.contains("excel-left-num")) {
        document.querySelectorAll(".selected").forEach((ele) => {
          ele.classList.remove("selected");
          ele.classList.remove("selected-col");
          ele.classList.remove("selected-row");
        });
        target.classList.add("selected");
        var rowIndex = target.getAttribute("row-index");
        document
          .querySelectorAll(".excel-cell.excel-cell-row-" + rowIndex)
          .forEach((ele) => {
            ele.classList.add("selected");
            ele.classList.add("selected-row");
          });

        this.$emit("on-row-select", rowIndex);
      } else if (target.classList.contains("excel-cell")) {
        document.querySelectorAll(".selected").forEach((ele) => {
          ele.classList.remove("selected");
          ele.classList.remove("selected-col");
          ele.classList.remove("selected-row");
        });

        target.classList.add("selected");
        var rowIndex = target.getAttribute("row-index");
        var colIndex = target.getAttribute("col-index");

        this.$emit("on-cell-select", rowIndex, colIndex, target.innerText);
      }
    },
    openExcelFile(file) {
      this.$emit('on-before-open')
      var that = this;
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = e.target.result;
        console.info("start read workbook");
        var workbook = XLSX.read(data, {
          type: "binary",
        });
        var sheetName = workbook.SheetNames[0];
        if (sheetName) {
          renderWorkbookSheet(workbook, sheetName);
          that.$emit('on-after-open')
        }
      };
      reader.readAsBinaryString(file);
    },
    openExcelData(data) {
      this.$emit('on-before-open')
      var that = this;
      var workbook = XLSX.read(data, {
        type: "binary",
      });
      var sheetName = workbook.SheetNames[0];
      if (sheetName) {
        renderWorkbookSheet(workbook, sheetName);
        that.$emit('on-after-open')
      }
    },
    reachBottom() {
      this.$emit("on-reach-bottom");
    },
  },
};
</script>

<style lang="scss">
.excel-table {
  font-size: 1.2em;
  border-collapse: collapse;

  :hover {
    cursor: pointer;
  }

  thead {
    th.excel-head-th {
      border-left: 1px solid  #d4d4d4;
      border-right: 1px solid  #d4d4d4;
      background-color: #e8e8e8;
    }
    th.selected {
      border-bottom: 2px solid green;
    }
  }

  .excel-row {
    height: 30px;

    .excel-left-num {
      border-top: 1px solid  #d4d4d4;
      border-bottom: 1px solid  #d4d4d4;
      background-color: #e8e8e8;
    }
    .excel-left-num.selected {
      border-right: 2px solid green;
    }

    .excel-cell {
      border: 1px solid #d4d4d4;
    }

    .excel-cell.selected-row {
      border-top: 2px solid green;
      border-bottom: 2px solid green;
      background-color: #e8e8e8;
    }
    .excel-cell.selected-col {
      border-left: 2px solid green;
      border-right: 2px solid green;
      background-color: #e8e8e8;
    }

    .excel-cell.selected:not(.selected-row):not(.selected-col) {
      border: 2px solid green;
    }

    .excel-cell.active {
      border: 2px solid green;
    }
  }
}
</style>