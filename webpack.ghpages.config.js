const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = {
  mode: "production",
  devtool: "source-map",
  entry: {
    polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
    taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
    commands: "./src/commands/commands.js"
  },
  resolve: {
    extensions: [".js", ".html"]
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: "babel-loader"
      },
      {
        test: /\.html$/,
        exclude: /node_modules/,
        use: "html-loader"
      },
      {
        test: /\.(png|jpg|jpeg|gif|ico)$/,
        type: "asset/resource",
        generator: {
          filename: "assets/[name][ext]"
        }
      }
    ]
  },
  output: {
    filename: "[name].bundle.js",
    path: `${__dirname}/dist-ghpages`,
    clean: true
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["polyfill", "taskpane"]
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["polyfill", "commands"]
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "manifest*.xml",
          to: "[name][ext]"
        },
        {
          from: "assets",
          to: "assets"
        }
      ]
    })
  ]
};
