var path = require("path");
var webpack = require("webpack");
const UglifyJSPlugin = require('uglifyjs-webpack-plugin');
var CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    target: "web",
    entry: {
        App: "./scripts/App.tsx"
    },
    output: {
        filename: "scripts/[name].js",
        libraryTarget: "amd"
    },
    externals: [
        {
            "q": true,
            "react": true,
            "react-dom": true
        },
        /^VSS\/.*/, /^TFS\/.*/, /^q$/
    ],
    resolve: {
        extensions: [".webpack.js", ".web.js", ".ts", ".tsx", ".js"],
        moduleExtensions: ["-loader"],
        alias: { 
            "OfficeFabric": path.resolve(__dirname, "node_modules/office-ui-fabric-react/lib-amd"),
            "VSTS_Extension": path.resolve(__dirname, "node_modules/vsts-extension-react-widgets/lib-amd")
        }
    },
    module: {
        loaders: [
            {
                test: /\.tsx?$/,
                loader: "ts-loader"
            },
            {
                test: /\.s?css$/,
                loaders: ["style-loader", "css-loader", "sass-loader"]
            }
        ]
    },
    plugins: [
        new webpack.DefinePlugin({
            "process.env.NODE_ENV": JSON.stringify("production")
        }),
        new UglifyJSPlugin({
            minimize: true,
            compress: {
                warnings: false
            },
            output: {
                comments: false
            }
        }),
        new CopyWebpackPlugin([
            { from: "./node_modules/vss-web-extension-sdk/lib/VSS.SDK.min.js", to: "scripts/libs/VSS.SDK.min.js" },
            { from: "./node_modules/es6-promise/dist/es6-promise.min.js", to: "scripts/libs/es6-promise.min.js" },           
            
            { from: "./node_modules/office-ui-fabric-react/dist/css/fabric.min.css", to: "css/libs/fabric.min.css" },

            { from: "./img", to: "img" },
            { from: "./index.html", to: "./" },
            { from: "./README.md", to: "README.md" },
            { from: "./vss-extension.json", to: "vss-extension.json" }
        ])
    ]
}