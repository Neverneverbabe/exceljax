const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');

module.exports = {
  entry: {
    taskpane: './src/taskpane/taskpane.js',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: 'taskpane.html',
      template: './src/taskpane/taskpane.html',
      chunks: ['taskpane'],
    }),
    new webpack.DefinePlugin({
      LM_API_URL: JSON.stringify(process.env.LM_API_URL || 'http://localhost:1234/v1/models'),
    }),
  ],
  module: {
    rules: [
      {
        test: /\.css$/,
        use: [
          {
            loader: 'file-loader',
            options: {
              name: '[name].[ext]',
            },
          },
        ],
      },
    ],
  },
  mode: 'production',
  devtool: 'source-map',
};
