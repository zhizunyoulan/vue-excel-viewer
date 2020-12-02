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
            .rule("svg")
            .exclude.add(resolve("src/icons"))
            .end();
        config.module
            .rule("icons")
            .test(/\.svg$/)
            .include.add(resolve("src/icons"))
            .end()
            .use("svg-sprite-loader")
            .loader("svg-sprite-loader")
            .options({
                symbolId: "icon-[name]",
            })
            .end();

        config.when(process.env.NODE_ENV == "development", (config) => {
            config
                .plugin("ScriptExtHtmlWebpackPlugin")
                .after("html")
                .use("script-ext-html-webpack-plugin", [{
                    // `runtime` must same as runtimeChunk name. default is `runtime`
                    inline: /runtime\..*\.js$/,
                }, ])
                .end();
        });
    },
};