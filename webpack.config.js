// webpack.config.js
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");

const helpers = require("./config/helpers"); // Assuming you have a helpers.js in config folder

module.exports = {
  mode: "production", // Or "development" for easier debugging
  entry: {
    polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
    taskpane: "./src/taskpane/taskpane.js",
    commands: "./src/commands/commands.js",
    login: "./src/login/login.js", // Assuming you have a login entry point
  },
  output: {
    clean: true,
    filename: "[name].js",
    path: helpers.root("dist"), // Output to dist folder
    publicPath: "/excel-addin-demo/", // <--- IMPORTANT: Update this to your GitHub repo name!
  },
  resolve: {
    extensions: [".js", ".jsx", ".json", ".css", ".html"],
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: "babel-loader",
          options: {
            presets: ["@babel/preset-env"],
          },
        },
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"],
      },
      {
        test: /\.html$/,
        exclude: /node_modules/,
        use: "html-loader",
      },
      {
        test: /\.(png|jpg|jpeg|gif|svg|woff|woff2|ttf|eot)$/,
        type: "asset/resource",
        generator: {
          filename: "assets/[name][ext]",
        },
      },
    ],
  },
  plugins: [
    new CleanWebpackPlugin(),
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane", "polyfill"],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"],
    }),
    new HtmlWebpackPlugin({
      filename: "login.html", // Assuming you have a login.html
      template: "./src/login/login.html",
      chunks: ["login", "polyfill"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "assets/*",
          to: "assets/[name][ext]",
        },
        {
          from: "manifest.xml",
          to: "manifest.xml",
        },
        {
          from: "functions.json",
          to: "functions.json",
        },
      ],
    }),
    new webpack.ProvidePlugin({
      Promise: ["es6-promise", "Promise"],
    }),
  ],
  devServer: {
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    hot: true,
    static: {
      directory: helpers.root("dist"),
      publicPath: "/excel-addin-demo/", // <--- IMPORTANT: Also set for dev server
    },
    port: 3000,
  },
};


// /* eslint-disable no-undef */

// const devCerts = require("office-addin-dev-certs");
// const CopyWebpackPlugin = require("copy-webpack-plugin");
// const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
// const HtmlWebpackPlugin = require("html-webpack-plugin");
// const path = require("path");

// const urlDev = "https://localhost:3000/";
// const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

// /* global require, module, process, __dirname */

// async function getHttpsOptions() {
//   const httpsOptions = await devCerts.getHttpsServerOptions();
//   return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
// }

// module.exports = async (env, options) => {
//   const dev = options.mode === "development";
//   const config = {
//     devtool: "source-map",
//     entry: {
//       polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
//       taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
//       commands: "./src/commands/commands.js",
//       functions: "./src/functions/functions.js",
//     },
//     output: {
//       clean: true,
//     },
//     resolve: {
//       extensions: [".html", ".js"],
//     },
//     module: {
//       rules: [
//         {
//           test: /\.js$/,
//           exclude: /node_modules/,
//           use: {
//             loader: "babel-loader",
//           },
//         },
//         {
//           test: /\.html$/,
//           exclude: /node_modules/,
//           use: "html-loader",
//         },
//         {
//           test: /\.(png|jpg|jpeg|gif|ico)$/,
//           type: "asset/resource",
//           generator: {
//             filename: "assets/[name][ext][query]",
//           },
//         },
//       ],
//     },
//     plugins: [
//       new CustomFunctionsMetadataPlugin({
//         output: "functions.json",
//         input: "./src/functions/functions.js",
//       }),
//       new HtmlWebpackPlugin({
//         filename: "functions.html",
//         template: "./src/functions/functions.html",
//         chunks: ["polyfill", "functions"],
//       }),
//       new HtmlWebpackPlugin({
//         filename: "taskpane.html",
//         template: "./src/taskpane/taskpane.html",
//         chunks: ["polyfill", "taskpane"],
//       }),
//       new CopyWebpackPlugin({
//         patterns: [
//           {
//             from: "assets/*",
//             to: "assets/[name][ext][query]",
//           },
//           {
//             from: "manifest*.xml",
//             to: "[name]" + "[ext]",
//             transform(content) {
//               if (dev) {
//                 return content;
//               } else {
//                 return content.toString().replace(new RegExp(urlDev + "(?:public/)?", "g"), urlProd);
//               }
//             },
//           },
//         ],
//       }),
//       new HtmlWebpackPlugin({
//         filename: "commands.html",
//         template: "./src/commands/commands.html",
//         chunks: ["polyfill", "commands"],
//       }),
//     ],
//     devServer: {
//       static: {
//         directory: path.join(__dirname, "dist"),
//         publicPath: "/public",
//       },
//       headers: {
//         "Access-Control-Allow-Origin": "*",
//       },
//       server: {
//         type: "https",
//         options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
//       },
//       port: process.env.npm_package_config_dev_server_port || 3000,
//     },
//   };

//   return config;
// };
