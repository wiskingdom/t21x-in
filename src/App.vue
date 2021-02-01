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
          label="검색(제)"
          style="margin-left: 10px"
          @click="externPop('headLine')"
        />
        <q-btn
          class="text-bold text-caption"
          color="green-8"
          unelevated
          label="검색(본)"
          style="margin-left: 10px"
          @click="externPop('text')"
        />
        <q-btn-dropdown
          class="text-bold text-caption"
          color="accent"
          unelevated
          :label="`설정(${newsTypeKor})`"
          style="margin-left: 10px"
        >
          <q-item flat clickable v-close-popup @click="initSet()"
            >(0) 서식 설정</q-item
          >
          <q-item clickable v-close-popup @click="dup1Set()"
            >(1-1) 중복 검토 1</q-item
          >
          <q-item clickable v-close-popup @click="dup2Set()"
            >(1-2) 중복 검토 2</q-item
          >
          <q-item clickable v-close-popup @click="eSet()">(2) 사설 검토</q-item>
          <q-item clickable v-close-popup @click="nSet()"
            >(3) 비기사 검토</q-item
          >
          <q-item clickable v-close-popup @click="ySet()"
            >(4) 일반기사 검토</q-item
          >
        </q-btn-dropdown>
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
              <td class="text-left">{{ record.DupID }}</td>
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
      <q-dialog v-model="alert">
        <q-card>
          <q-card-section class="bg-accent text-white">
            <div class="text-h6">{{ alertContent.title }}</div>
          </q-card-section>

          <q-card-section class="q-pt-none">
            <br />
            {{ alertContent.message }}
          </q-card-section>

          <q-card-actions align="right">
            <q-btn flat label="확인" color="primary" v-close-popup />
          </q-card-actions>
        </q-card>
      </q-dialog>
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
const newsTypeMap = {
  cho: "조",
  dong: "동",
  joong: "중",
  han: "한"
};
export default {
  name: "App",
  data() {
    return {
      alert: false,
      alertContent: { title: "", message: "" },
      newsType: "type",
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
      return markMap[this.record.Mark] || "미수정";
    },
    preMarkColor() {
      return markColorMap[this.record.PreMark] || "text-grey-9";
    },
    markColor() {
      return markColorMap[this.record.Mark] || "text-grey-9";
    },
    newsTypeKor() {
      return newsTypeMap[this.newsType] || "오류";
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
    popAlert({ title, message }) {
      this.alertContent = { title, message };
      this.alert = true;
    },
    async initSet() {
      await this.tryCatch(this.filter());
      await this.tryCatch(this.requireApprovedTag);
      await this.tryCatch(this.setRowColorGrid);
      await this.tryCatch(this.freezeFirstRow);
      await this.tryCatch(this.sort(0));
      this.popAlert({
        title: "설정 완료",
        message: "(0) 서식 설정을 완료하였습니다."
      });
    },
    async dup1Set() {
      await this.tryCatch(this.filter(8, "dup-1"));
      await this.tryCatch(this.sort(0));
      await this.tryCatch(this.sort(10));
      this.popAlert({
        title: "설정 완료",
        message: "(1-1) 중복 유형1 검토 설정을 완료하였습니다."
      });
    },
    async dup2Set() {
      await this.tryCatch(this.filter(8, "dup-2"));
      await this.tryCatch(this.sort(0));
      await this.tryCatch(this.sort(10));
      this.popAlert({
        title: "설정 완료",
        message: "(1-2) 중복 유형2 검토 설정을 완료하였습니다."
      });
    },
    async eSet() {
      await this.tryCatch(this.filter(8, "no-dup"));
      await this.tryCatch(this.filter(9, "e", true));
      await this.tryCatch(this.sort(12));
      this.popAlert({
        title: "설정 완료",
        message: "(2) 사설 검토 설정을 완료하였습니다."
      });
    },
    async nSet() {
      await this.tryCatch(this.filter(8, "no-dup"));
      await this.tryCatch(this.filter(9, "n", true));
      await this.tryCatch(this.sort(12));
      this.popAlert({
        title: "설정 완료",
        message: "(3) 비기사 검토 설정을 완료하였습니다."
      });
    },
    async ySet() {
      await this.tryCatch(this.filter(8, "no-dup"));
      await this.tryCatch(this.filter(9, "y", true));
      await this.tryCatch(this.sort(12));
      this.popAlert({
        title: "설정 완료",
        message: "(4) 일반기사 검토 설정을 완료하였습니다."
      });
    },
    externPop(mode) {
      const linkPre = this.record.SearchLink.match(/^.+(query|keyword)=/)[0];
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
          Dup,
          PreMark,
          DupID,
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
          Dup,
          PreMark,
          DupID,
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
    /* setting actions */

    filter(colNum, value, notClear) {
      return async function() {
        await window.Excel.run(async context => {
          const sheet = context.workbook.worksheets.getItem("data");
          const usedRange = sheet.getUsedRange();
          notClear || (await sheet.autoFilter.clearCriteria());
          await sheet.autoFilter.apply(usedRange, colNum, {
            values: [value],
            filterOn: window.Excel.FilterOn.values
          });
          await context.sync();
        });
      };
    },

    sort(colNum) {
      return async function() {
        await window.Excel.run(async context => {
          const sheet = context.workbook.worksheets.getItem("data");
          const usedRange = sheet.getUsedRange();
          const usedAddress = usedRange.load("address");
          await context.sync();
          const dataRange = sheet.getRange(
            usedAddress.address.replace(/A1/, "A2")
          );

          // sort the table by the "Amount" column
          const sortFields = [
            {
              key: colNum,
              ascending: true
            }
          ];
          dataRange.sort.apply(sortFields);

          await context.sync();
        });
      };
    },
    async requireApprovedTag() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const markRange = sheet.getRange("L:L");
        const protectedRange = sheet.getRanges("A:K, M:S");

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

    async setRowColorGrid() {
      await window.Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem("data");
        const range = sheet.getRange("L:L");

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
    async getNewsType() {
      await window.Excel.run(async context => {
        const workbookName = context.workbook.load("name");
        await context.sync();
        this.newsType = workbookName.name.split(/\d+/)[0];
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
    this.tryCatch(this.getNewsType);
    this.getValues(this.rowAddress);
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
