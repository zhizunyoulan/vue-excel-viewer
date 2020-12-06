import Vue from 'vue'
import App from './App'

import VueExcelViewer from '@/index.js'
Vue.use(VueExcelViewer)

Vue.config.productionTip = false

new Vue({
    el: "#app",
    components: {
        App,
    },
    template: "<App/>",
    render: (h) => h(App),
});