const path = require('path');

module.exports = {
  entry: {
    taskpane: './src/taskpane.js',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
  },
  module: {
    rules: [
      {
        test: /\.html$/,
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
  // Add mode and devtool for production clarity
  mode: 'production',
  devtool: 'source-map',
};
