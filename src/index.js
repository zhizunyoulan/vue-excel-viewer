const components = []
const _getModuleFrom = (root, path) => require('@/' + root + '/' + path + '.vue')

    //components
const componentPaths = require.context('@/components', true, /\.vue$/).keys().map(i => {
    return i.match(/\.\/(.*)\.vue/)[1]
})
componentPaths.forEach((path) => {
    const module = _getModuleFrom('components', path)
    if (module.default && module.default.automount) {
        console.info('automount module:', module)
        components.push(module.default)
    }
});
// 定义 install 方法
const install = function(Vue, option) {
    if (install.installed) return;
    install.installed = true;
        // 遍历并注册全局组件
    components.map((component) => {
        Vue.component(component.name, component);
    });

};

if (typeof window !== "undefined" && window.Vue) {
    install(window.Vue);
}

export default {
    // 导出的对象必须具备一个 install 方法
    install,
    // 组件列表
    ...components,
};