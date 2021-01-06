import Vue from "vue";
import App from "./App.vue";
import "./quasar";

Vue.config.productionTip = false;

window.Office.initialize = () => {
  new Vue({
    render: h => h(App)
  }).$mount("#app");
};
