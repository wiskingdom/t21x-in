<template>
  <q-layout view="hHh lpR fFf">
    <q-header bordered class="bg-grey-2 text-primary" height-hint="98">
      <!--
      <q-toolbar>
        
        <q-toolbar-title>
          <q-avatar>
            <img src="https://cdn.quasar.dev/logo/svg/quasar-logo.svg" />
          </q-avatar>
          Title
        </q-toolbar-title>
        
      </q-toolbar>
      -->
      <q-tabs align="left" class="padding">
        <span style="width:140px;display:inline-block;"
          >최종분류: {{ mark }}</span
        >
        <span>자동분류: {{ preMark }}</span>
        <q-space />
        <q-route-tab to="/page3" label="Page Three" />
      </q-tabs>
    </q-header>

    <q-page-container>
      <div id="app">
        <div class="record-headline">
          <div class="padding">
            <p>{{ record.HeadLine }}</p>
          </div>
        </div>
        <div class="padding">
          <p>{{ address }}</p>
          <p>{{ rowAddress }}</p>
          <p v-for="(item, index) in newsPs" v-bind:key="`p${index}`">
            {{ item }}
          </p>
        </div>
      </div>
    </q-page-container>
  </q-layout>
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
.padding {
  padding: 10px;
}
</style>
