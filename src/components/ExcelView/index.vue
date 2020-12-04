<template>
  <div class="excel-panel">
    <el-scrollbar
      ref="excelPanel"
      class="excel-panel"
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
    </el-scrollbar>
  </div>
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

function renderSheetAt(domElement, maxColRange, dataArray, merges, firstRowIndex) {
  var tableHeadElement = domElement.querySelector("thead");
  var emptyTh = document.createElement("th");
  tableHeadElement.appendChild(emptyTh);
  for (var i = 0; i < maxColRange; i++) {
    var colNum = i + 1;
    var sign = parseNumToChars(i);
    var th = document.createElement("th");
    th.classList.add("excel-head-th");
    th.classList.add("excel-cell-col-" + colNum);
    th.setAttribute("col-num", colNum);
    th.setAttribute("col-sign", sign);
    th.innerText = sign;
    tableHeadElement.appendChild(th);
  }
  var tableBodyElement = domElement.querySelector("tbody");
  var invalidCells = {};
  if(isNaN(firstRowIndex) || firstRowIndex < 0) {
    firstRowIndex = 0;
  }
  // console.info('firstRowIndex',firstRowIndex)
  for (var ri = firstRowIndex; ri < dataArray.length; ri++) {
    var rowNum = ri + 1;
    var row = dataArray[ri];
    var tr = document.createElement("tr");
    tr.setAttribute("row-num", rowNum);
    tr.classList.add("excel-row");
    var emptyTd = document.createElement("td");
    emptyTd.classList.add("excel-left-num");
    
    emptyTd.classList.add("excel-cell-row-" + rowNum);

    emptyTd.setAttribute("row-num", rowNum);
    emptyTd.innerText = rowNum;
    tr.appendChild(emptyTd);
    for (var ci = 0; ci < maxColRange; ci++) {
      var colNum = ci + 1;
      var cellKey = ri + "-" + ci;
      // debugger
      // console.info(cellKey, invalidCells[cellKey]);
      if (!invalidCells[cellKey]) {
        var td = document.createElement("td");
        td.classList.add("excel-cell");
        td.setAttribute("row-num", rowNum);
        td.setAttribute("col-num", colNum);
        td.classList.add("excel-cell-col-" + colNum);
        td.classList.add("excel-cell-row-" + rowNum);
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

function renderWorkbookSheet(workbook, sheetName, firstRowIndex, minColCounts) {
  var worksheet = workbook.Sheets[sheetName];
  // console.info("renderWorkbookSheet", worksheet);
  var defaultRange = worksheet["!ref"];
  var lastCellPosition = defaultRange.split(":")[1];
  var lastCellPositionMatchInfo = lastCellPosition.match(/(\D+)(\d+)/);
  // console.info("lastCellPositionMatchInfo", lastCellPositionMatchInfo);
  var merges = worksheet["!merges"];
  var sheetDatas = XLSX.utils.sheet_to_json(worksheet, {
    raw: false,
    header: 1,
  });
  // console.info("workbook sheet data", sheetDatas);
  var tableElement = document.querySelector(".excel-table");

  var maxColRange = parseCharsTo10(lastCellPositionMatchInfo[1]);
  if(!isNaN(minColCounts) && minColCounts > maxColRange) {
    maxColRange = minColCounts;
  }
  // console.info("maxColRange", maxColRange);
  renderSheetAt(tableElement, maxColRange, sheetDatas, transformMerges(merges), firstRowIndex);
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
  name: "ExcelView",
  automount: true,
  data() {
    return {
      isOpened: false,
      isScrollAtTop: true,
      isScrollAtBottom: false,
    };
  },
  props: {
    height: {
      type: Number,
      default: 500,
    },
    firstRowIndex: {
      type: Number
    },
    minColCounts: {
      type: Number
    }
  },
  mounted() {
    var that = this;
    document.querySelector(
      ".excel-panel .el-scrollbar__wrap"
    ).onscroll = function (e) {
      var scrollTop = e.target.scrollTop;

      if (scrollTop == 0) {
        if (!this.isScrollAtTop) {
          that.$emit("on-reach-top");
        }
        this.isScrollAtTop = true;
      } else {
        this.isScrollAtTop = false;

        var scrollTopMax = e.target.scrollTopMax;
        if (scrollTop == scrollTopMax) {
          if (!this.isScrollAtBottom) {
            that.$emit("on-reach-bottom");
          }
          this.isScrollAtBottom = true;
        } else {
          this.isScrollAtBottom = false;
        }
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
        var colNum = target.getAttribute("col-num");
        document
          .querySelectorAll(".excel-cell.excel-cell-col-" + colNum)
          .forEach((ele) => {
            ele.classList.add("selected");
            ele.classList.add("selected-col");
          });
        this.$emit("on-col-select", colNum);
      } else if (target.classList.contains("excel-left-num")) {
        document.querySelectorAll(".selected").forEach((ele) => {
          ele.classList.remove("selected");
          ele.classList.remove("selected-col");
          ele.classList.remove("selected-row");
        });
        target.classList.add("selected");
        var rowNum = target.getAttribute("row-num");

        var selectRowValues = [];
        document
          .querySelectorAll(".excel-cell.excel-cell-row-" + rowNum)
          .forEach((ele) => {
            var colNum = ele.getAttribute("col-num");
            if (!isNaN(colNum)) {
              selectRowValues[colNum] = ele.innerText;
            }
            ele.classList.add("selected");
            ele.classList.add("selected-row");
          });

        this.$emit("on-row-select", rowNum, selectRowValues);
      } else if (target.classList.contains("excel-cell")) {
        document.querySelectorAll(".selected").forEach((ele) => {
          ele.classList.remove("selected");
          ele.classList.remove("selected-col");
          ele.classList.remove("selected-row");
        });

        target.classList.add("selected");
        var rowNum = target.getAttribute("row-num");
        var colNum = target.getAttribute("col-num");

        document
          .querySelector("th.excel-cell-col-" + colNum)
          .classList.add("selected");
        document
          .querySelector("td.excel-cell-row-" + rowNum)
          .classList.add("selected");

        this.$emit("on-cell-select", rowNum, colNum, target.innerText);
      }
    },
    openExcelFile(file) {
      if (!this.isOpened) {
        this.isOpened = true;

        this.$emit("on-before-open");
        var that = this;
        const reader = new FileReader();
        reader.onload = function (e) {
          const data = e.target.result;
          // console.info("start read workbook");
          var workbook = XLSX.read(data, {
            type: "binary",
          });
          var sheetName = workbook.SheetNames[0];
          if (sheetName) {
            renderWorkbookSheet(workbook, sheetName, that.firstRowIndex, that.minColCounts);
            that.$emit("on-after-open");
          }
        };
        reader.readAsBinaryString(file);
      } else {
        throw "excel-view 不能重复打开";
      }
    },
    openExcelData(data,) {
      if (!this.isOpened) {
        this.isOpened = true;
        this.$emit("on-before-open");
        var that = this;
        var workbook = XLSX.read(data, {
          type: "binary",
        });
        var sheetName = workbook.SheetNames[0];
        if (sheetName) {
          renderWorkbookSheet(workbook, sheetName, that.firstRowIndex, that.minColCounts);
          that.$emit("on-after-open");
        }
      } else {
        throw "excel-view 不能重复打开";
      }
    },
  },
};
</script>

<style lang="scss">
.excel-panel .excel-table {
  // font-size: 1.2em;
  border-collapse: collapse;
  user-select: none;

  :hover {
    cursor: pointer;
  }

  thead {
    th.excel-head-th {
      position: sticky;
      top: 0;
      border-left: 1px solid #bbbbbb;
      border-right: 1px solid #bbbbbb;
      background-color: #e8e8e8;
      background-clip: padding-box;
    }

    th.excel-head-th.selected {
      background-color: #d6d6d6;
    }

    th.excel-head-th.selected::after {
      content: "";
      background-color: #42a642;
      position: absolute;
      left: 0;
      bottom: 0;
      width: 100%;
      height: 2px;
    }
  }

  .excel-row {
    height: 30px;

    .excel-left-num {
      position: sticky;
      left: 0;
      border-top: 1px solid #d4d4d4;
      border-bottom: 1px solid #d4d4d4;
      background-color: #e8e8e8;
      background-clip: padding-box;
      text-align: center;
      justify-content: space-around;
      align-items: center;
      padding-left: 5px;
      padding-right: 5px;
    }

    .excel-left-num.selected {
      background-color: #d6d6d6;
    }

    .excel-left-num.selected::after {
      content: "";
      background-color: #42a642;
      position: absolute;
      top: 0;
      right: 0;
      height: 100%;
      width: 2px;
    }

    .excel-cell {
      border: 1px solid #d4d4d4;
      white-space: nowrap;
      width: 50px;
    }

    .excel-cell.selected-row {
      border-top: 2px solid #42a642;
      border-bottom: 2px solid #42a642;
      background-color: #e8e8e8;
    }
    .excel-cell.selected-col {
      border-left: 2px solid #42a642;
      border-right: 2px solid #42a642;
      background-color: #e8e8e8;
    }

    .excel-cell.selected:not(.selected-row):not(.selected-col) {
      border: 2px solid #42a642;
    }

    .excel-cell.active {
      border: 2px solid #42a642;
    }
  }
}
</style>