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
        <span style="width:140px;display:inline-block;" :class="markColor"
          ><strong>최종분류: {{ markKor }}</strong></span
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
      <q-tabs align="left" class="text-grey-9 padding">
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
  p: "일반기사",
  n: "비기사",
  s: "사설",
  d: "중복기사"
};
const markColorMap = {
  p: "text-deep-purple-9",
  n: "text-blue-grey-9",
  s: "text-amber-10",
  d: "text-red-9"
};
export default {
  name: "App",
  data() {
    return {
      address: "A1",
      rowAddress: "1",
      record: {}
    };
  },
  computed: {
    preMarkKor() {
      return markMap[this.record.PreMark] || "";
    },
    markKor() {
      const finalMark = this.record.Mark || this.record.PreMark;
      return markMap[finalMark] || "";
    },
    preMarkColor() {
      return markColorMap[this.record.PreMark] || "text-grey-9";
    },
    markColor() {
      return this.record.Mark
        ? markColorMap[this.record.Mark]
        : markColorMap[this.record.PreMark];
    },
    newsPs() {
      return this.record.NewsText ? this.record.NewsText.split(/<LFCR>/) : [];
    }
  },
  watch: {
    address(newAddress) {
      this.rowAddress = this.getRowAddress(newAddress);
    },
    rowAddress(newAddress, oldAddress) {
      newAddress === "1" || this.getValues(newAddress);
      this.setRowColor(newAddress, "yellow");
      this.setRowColor(oldAddress, null);
    }
  },
  methods: {
    externPop() {
      const link = this.record.SearchLink;
      const linkUri = encodeURI(link);
      window.open(linkUri, "popup");
      return false;
    },
    getRowAddress(address) {
      return address.match(/\d+/)[0];
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
    },
    markValidate() {
      this.tryCatch(this.requireApprovedTag);
    },
    async requireApprovedTag() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const markRange = sheet.getRange("N:N");

        markRange.dataValidation.clear();

        let approvedListRule = {
          list: {
            inCellDropDown: true,
            source: "p,n,s"
          }
        };
        markRange.dataValidation.rule = approvedListRule;

        await context.sync();
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
