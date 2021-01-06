<template>
  <div id="app">
    <div class="record-meta">
      <div class="padding">
        <table>
          <tr>
            <td style="width:150px;">
              <strong>최종분류: {{ mark }}</strong>
            </td>
            <td>자동분류: {{ preMark }}</td>
          </tr>
        </table>
      </div>
    </div>
    <div class="record-headline">
      <div class="padding">
        <p>{{ record.HeadLine }}</p>
      </div>
    </div>
    <div class="record-text">
      <p>{{ address }}</p>
      <p>{{ rowAddress }}</p>
      <p v-for="(item, index) in newsPs" v-bind:key="`p${index}`">{{ item }}</p>
    </div>
  </div>
</template>

<script>
const markMap = {
  p: "일반기사",
  n: "비기사",
  s: "사설",
  d: "중복기사"
};
export default {
  name: "App",
  data() {
    return {
      address: "init",
      rowAddress: "init",
      record: {}
    };
  },
  computed: {
    preMark() {
      return markMap[this.record.PreMark] || "";
    },
    mark() {
      const finalMark = this.record.Mark || this.record.PreMark;
      return markMap[finalMark] || "";
    },
    newsPs() {
      return this.record.NewsText ? this.record.NewsText.split(/<LFCR>/) : [];
    }
  },
  watch: {
    address(newAddress) {
      this.rowAddress = this.getRowAddress(newAddress);
    },
    rowAddress(newAddress) {
      this.getValues(newAddress);
    }
  },
  methods: {
    getRowAddress(address) {
      return address.match(/\d+/)[0];
    },
    onSetColor() {
      window.Excel.run(async context => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = "green";
        await context.sync();
      });
    },
    connect() {
      this.tryCatch(this.registerEventHandlers);
    },
    async tryCatch(callback) {
      try {
        await callback();
      } catch (error) {
        console.error(error);
      }
    },
    async registerEventHandlers() {
      await window.Excel.run(async context => {
        let sheet = context.workbook.worksheets.getItem("data");
        sheet.onSelectionChanged.add(this.onWorksheetSelectionChange);

        await context.sync();
      });
    },
    async onWorksheetSelectionChange(args) {
      await window.Excel.run(async context => {
        this.updateAddress(args.address);
        await context.sync();
      });
    },
    updateAddress(payload) {
      this.address = payload;
    },
    async getValues(rowAddress) {
      const allRowRange = `A${rowAddress}:S${rowAddress}`; // 레코드 열 범위
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const range = sheet.getRange(allRowRange);
        range.load("values");

        await context.sync();
        const [
          ID,
          NewsId,
          PageType,
          PrintingPage,
          SectionPage,
          SubjectCode,
          T21Class,
          DateLine,
          PreMark,
          HeadLine,
          SubHeadLine,
          ByLine,
          Dup,
          Mark,
          NewsText,
          SearchLink,
          WordCount
        ] = range.values[0];
        this.record = {
          ID,
          NewsId,
          PageType,
          PrintingPage,
          SectionPage,
          SubjectCode,
          T21Class,
          DateLine,
          PreMark,
          HeadLine,
          SubHeadLine,
          ByLine,
          Dup,
          Mark,
          NewsText,
          SearchLink,
          WordCount
        };
      });
    }
  },
  created() {
    this.connect();
  }
};
</script>

<style>
.record-meta {
  background: #d1e3f0;
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  height: 80px;
  bottom: 0;
  padding: 10px;
  font-size: 14px;
  overflow: auto;
}
.record-headline {
  background: #d1e3f0;
  position: fixed;
  top: 70px;
  left: 0;
  right: 0;
  height: 80px;
  bottom: 0;
  padding: 10px;
  font-size: 14px;
  overflow: auto;
}

.record-text {
  background: #fff;
  position: fixed;
  top: 160px;
  left: 0;
  right: 0;
  bottom: 0;
  padding: 10px;
  font-size: 14px;
  overflow: auto;
}
.padding {
  padding: 0px 10px;
}
</style>
