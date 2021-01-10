<template>
  <q-layout view="hHh lpR fFf">
    <q-header bordered class="bg-grey-2 text-primary" height-hint="98">
      <q-tabs align="left" class="padding">
        <span
          style="width:160px;display:inline-block;"
          class="text-subtitle1"
          :class="markColor"
          ><strong>작업자분류: {{ markKor }}</strong></span
        >
        <span style="width:250px;display:inline-block;" :class="preMarkColor"
          >추정분류: {{ preMarkKor }} ({{ record.PreWhy }})</span
        >
        <q-btn
          class="text-bold text-caption"
          color="green-8"
          unelevated
          label="기사검색(제)"
          style="margin-left: 10px"
          @click="externPop('headLine')"
        />
        <q-btn
          class="text-bold text-caption"
          color="green-8"
          unelevated
          label="기사검색(본)"
          style="margin-left: 10px"
          @click="externPop('text')"
        />
        <q-btn
          class="text-bold text-caption"
          color="accent"
          unelevated
          label="서식세팅"
          style="margin-left: 10px"
          @click="initForm()"
        />
      </q-tabs>
      <q-tabs align="left" class="text-grey-9 padding">
        <q-markup-table flat bordered>
          <thead>
            <tr>
              <th class="text-left">ID</th>
              <th class="text-left">중복ID</th>
              <th class="text-left">날짜</th>
              <th class="text-left">작성자</th>
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
              <td class="text-left">{{ record.ByLine }}</td>
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

    <q-footer bordered class="bg-white text-accent">
      <div class="padding">
        <p v-for="(item, index) in lastString" v-bind:key="`l${index}`">
          {{ item }}
        </p>
      </div>
    </q-footer>
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
      event: {},
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
    },
    lastString() {
      const length = 100;
      const text = this.record.NewsText;
      return text.length < length
        ? [text]
        : text
            .slice(0 - length)
            .replace(/^.+?\s+/, "")
            .split(/<LFCR>/);
    }
  },
  watch: {
    address(newAddress) {
      const rowAddresses = this.getRowAddresses(newAddress);
      this.rowAddress = rowAddresses.reverse()[0];
    },
    rowAddress(newAddress) {
      newAddress === "1" || this.getValues(newAddress);
    }
  },
  methods: {
    initForm() {
      this.tryCatch(this.requireApprovedTag);
      this.tryCatch(this.setRowColorGrid);
      this.tryCatch(this.setFilter);
      this.tryCatch(this.freezeFirstRow);
    },
    externPop(mode) {
      const linkPre = this.record.SearchLink.match(/^.+query=/)[0];
      const link =
        mode === "headLine"
          ? linkPre + this.record.HeadLine.replace(/[^가-힣\w]+/g, " ")
          : this.record.SearchLink.replace(/CRLF|LFCR/g, "");
      const linkUri = encodeURI(link);
      window.open(linkUri, "popup");
      return false;
    },
    getRowAddresses(address) {
      return address.match(/\d+/g);
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
          PreWhy,
          WordCount,
          SearchLink
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
          PreWhy,
          SearchLink,
          WordCount
        };
      });
    },
    /* initializer */
    async tryCatch(callback) {
      try {
        await callback();
      } catch (error) {
        console.error(error);
      }
    },
    async registerEventHandler() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        this.event = sheet.onSelectionChanged.add(
          this.onWorksheetSelectionChange
        );
        await context.sync();
      });
    },
    async onWorksheetSelectionChange(args) {
      await window.Excel.run(async context => {
        this.address = args.address;
        await context.sync();
      });
    },
    async requireApprovedTag() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const markRange = sheet.getRange("K:K");
        const protectedRange = sheet.getRanges("A:J, L:Q");

        markRange.dataValidation.clear();
        protectedRange.dataValidation.clear();

        let approvedListRule = {
          list: {
            inCellDropDown: true,
            source: "y,n,e,z"
          }
        };
        let protectedListRule = {
          list: {
            inCellDropDown: false,
            source: "protected"
          }
        };
        markRange.dataValidation.rule = approvedListRule;
        protectedRange.dataValidation.rule = protectedListRule;

        await context.sync();
      });
    },
    async freezeFirstRow() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        sheet.freezePanes.freezeRows(1);

        await context.sync();
      });
    },
    async setFilter() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const usedRange = sheet.getUsedRange();

        sheet.autoFilter.apply(usedRange);
        await context.sync();
      });
    },
    async setRowColorGrid() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const range = sheet.getRange("K:K");

        range.format.borders.getItem("InsideHorizontal").style = "Dot";
        range.format.borders.getItem("InsideVertical").style = "Dot";
        range.format.borders.getItem("EdgeBottom").style = "Dot";
        range.format.borders.getItem("EdgeLeft").style = "Dot";
        range.format.borders.getItem("EdgeRight").style = "Dot";
        range.format.borders.getItem("EdgeTop").style = "Dot";

        range.format.borders.getItem("InsideHorizontal").weight = "Hairline";
        range.format.borders.getItem("InsideVertical").weight = "Hairline";
        range.format.borders.getItem("EdgeBottom").weight = "Hairline";
        range.format.borders.getItem("EdgeLeft").weight = "Hairline";
        range.format.borders.getItem("EdgeRight").weight = "Hairline";
        range.format.borders.getItem("EdgeTop").weight = "Hairline";

        range.format.fill.color = "#F8FAF4";
        await context.sync();
      });
    },
    /* terminator */
    async removeEventHandler() {
      await window.Excel.run(async context => {
        this.event.remove();

        await context.sync();
      });
    }
  },
  created() {
    this.tryCatch(this.registerEventHandler);
  },
  beforeDestroy() {
    this.tryCatch(this.removeEventHandler);
  }
};
</script>

<style>
.padding {
  padding: 10px;
}
</style>
