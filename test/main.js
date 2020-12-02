import Vue from 'vue'
import App from './App'
import Element from 'element-ui'
import enLang from 'element-ui/lib/locale/lang/en'// 如果使用中文语言包请默认支持，无需额外引入，请删除该依赖

// import 'normalize.css/normalize.css' // A modern alternative to CSS resets
import 'element-ui/lib/theme-chalk/index.css';

Vue.use(Element)


Vue.config.productionTip = false

new Vue({
    el: "#app",
    components: {
        App,
    },
    template: "<App/>",
    render: (h) => h(App),
});