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
        <span
          style="width:160px;display:inline-block;"
          class="text-subtitle1"
          :class="markColor"
          ><strong>작업자분류: {{ markKor }}</strong></span
        >
        <span style="width:140px;display:inline-block;" :class="preMarkColor"
          >자동분류: {{ preMarkKor }}</span
        >
        <q-btn
          class="text-bold"
          color="green-8"
          unelevated
          label="기사검색"
          style="margin-left: 10px"
          @click="externPop()"
        />
      </q-tabs>
      <q-tabs align="left" class="text-grey-9 padding">
        <q-markup-table flat bordered>
          <thead>
            <tr>
              <th class="text-left">ID</th>
              <th class="text-left">중복ID</th>
              <th class="text-left">날짜</th>
              <th class="text-left">면종</th>
              <th class="text-left">페이지</th>
              <th class="text-left">단어수</th>
              <th class="text-left">주제</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td class="text-left">{{ record.ID }}</td>
              <td class="text-left">{{ record.Dup }}</td>
              <td class="text-left">{{ record.DateLine }}</td>
              <td class="text-left">{{ record.PageType }}</td>
              <td class="text-left">{{ record.PrintingPage }}</td>

              <td class="text-left">{{ record.WordCount }}</td>
              <td class="text-left">{{ record.SubjectCode }}</td>
            </tr>
          </tbody>
        </q-markup-table>
        <br />
      </q-tabs>
      <q-tabs align="left" class="text-grey-9 text-subtitle1 padding">
        <p>
          <strong>제목: {{ record.HeadLine }}</strong>
        </p>
      </q-tabs>
    </q-header>

    <q-page-container>
      <div id="app">
        <div class="padding">
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
  y: "일반기사",
  n: "비기사",
  e: "사설",
  d: "중복기사",
  z: "보류"
};
const markColorMap = {
  y: "text-primary",
  n: "text-orange-10",
  e: "text-accent",
  d: "text-negative",
  z: "text-amber-10"
};
export default {
  name: "App",
  data() {
    return {
      address: "A1",
      rowAddress: "1",
      colAddress: "A",
      record: {},
      log: ""
    };
  },
  computed: {
    preMarkKor() {
      return markMap[this.record.PreMark] || "";
    },
    markKor() {
      return markMap[this.record.Mark] || "미분류";
    },
    preMarkColor() {
      return markColorMap[this.record.PreMark] || "text-grey-9";
    },
    markColor() {
      return markColorMap[this.record.Mark] || "text-grey-9";
    },
    newsPs() {
      return this.record.NewsText ? this.record.NewsText.split(/<LFCR>/) : [];
    }
  },
  watch: {
    address(newAddress) {
      const rowAddresses = this.getRowAddresses(newAddress);
      const colAddresses = this.getColAddresses(newAddress);
      this.rowAddress = rowAddresses.reverse()[0];
      this.colAddress = colAddresses[0];

      rowAddresses.length === 1 &&
        this.setRowColor(rowAddresses.reverse()[0], "yellow");
      if (colAddresses.every(item => item === "K")) {
        this.unprotectDataSheet();
      } else {
        this.protectDataSheet();
      }
    },
    rowAddress(newAddress, oldAddress) {
      newAddress === "1" || this.getValues(newAddress);
      this.setRowColor(oldAddress, null);
      this.log = oldAddress;
    }
  },
  methods: {
    externPop() {
      const link = this.record.SearchLink.replace(/CRLF\+/g, "").replace(
        /\+/g,
        " "
      );
      const linkUri = encodeURI(link);
      window.open(linkUri, "popup");
      return false;
    },
    getRowAddresses(address) {
      return address.match(/\d+/g);
    },
    getColAddresses(address) {
      return address.match(/[A-Z]+/g);
    },
    setRowColor(rowAddress, color) {
      window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const rowRange = sheet.getRange(`A${rowAddress}:Q${rowAddress}`);
        if (color) {
          rowRange.format.fill.color = color;
        } else {
          rowRange.format.fill.clear();
        }

        await context.sync();
      });
    },
    bindX() {
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
        const sheet = context.workbook.worksheets.getItem("data");
        sheet.onSelectionChanged.add(this.onWorksheetSelectionChange);
        //sheet.onChanged.add(this.onValueChange);
        await context.sync();
      });
    },
    async onWorksheetSelectionChange(args) {
      await window.Excel.run(async context => {
        this.updateAddress(args.address);
        await context.sync();
      });
    },
    /*
    async onValueChange(args) {
      await window.Excel.run(async context => {
        this.changeValueAddress = args.address;
        await context.sync();
      });
    },
    */

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
          Dup,
          Mark,
          HeadLine,
          SubHeadLine,
          ByLine,
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
          Dup,
          Mark,
          HeadLine,
          SubHeadLine,
          ByLine,
          NewsText,
          SearchLink,
          WordCount
        };
      });
    },
    markValidate() {
      this.tryCatch(this.requireApprovedTag);
    },
    async requireApprovedTag() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const markRange = sheet.getRange("K:K");

        markRange.dataValidation.clear();

        let approvedListRule = {
          list: {
            inCellDropDown: true,
            source: "y,n,e,z"
          }
        };
        markRange.dataValidation.rule = approvedListRule;

        await context.sync();
      });
    },
    async protectDataSheet() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        sheet.load("protection/protected");

        await context.sync();

        if (!sheet.protection.protected) {
          const option = {
            allowAutoFilter: true,
            allowFormatRows: true,
            allowFormatCells: true,
            allowFormatColumns: true
          };
          sheet.protection.protect(option);
        }
      });
    },
    async unprotectDataSheet() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        sheet.protection.unprotect();
      });
    }
  },
  created() {
    this.bindX();
    this.markValidate();
  }
};
</script>

<style>
.padding {
  padding: 10px;
}
</style>
