const path = require("path");

function resolve(dir) {
    return path.join(__dirname, dir);
}

module.exports = {
    lintOnSave: false, //这里禁止使用eslint-loader
    pages: {
        index: {
            entry: "test/main.js",
            template: "public/index.html",
            filename: "index.html",
        },
    },
    configureWebpack: {
        resolve: {
            extensions: ["js", "vue"],
            alias: {
                "@": resolve("src"),
            },
        },
    },

    chainWebpack(config) {
        // set svg-sprite-loader
        config.module
            .rule('image')
            .test(/\.ico$/)
            .use('url-loader')
            .loader('url-loader')
            config.module
            .rule('image')
            .test(/\.cur$/)
            .use('url-loader')
            .loader('url-loader')
    },
};