const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const themeCss = fs.readFileSync(path.resolve(__dirname, "src/theme.css"), "utf8");

module.exports = async (_env, options) => {
  const dev = options.mode === "development";
  const https = dev
    ? await require("office-addin-dev-certs").getHttpsServerOptions()
    : undefined;

  return {
    entry: { taskpane: "./src/taskpane/taskpane.ts" },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: { extensions: [".ts", ".js"] },
    module: {
      rules: [{ test: /\.ts$/, use: "ts-loader", exclude: /node_modules/ }],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        themeCss,
      }),
      new HtmlWebpackPlugin({
        filename: "info.html",
        template: "./src/info/info.html",
        inject: false,
        themeCss,
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "src/theme.css", to: "theme.css" },
        ],
      }),
    ],
    devServer: {
      server: { type: "https", options: https },
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" },
    },
  };
};
