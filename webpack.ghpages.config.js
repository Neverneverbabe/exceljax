// webpack.ghpages.config.js

import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export default {
  entry: './src/taskpane/taskpane.js',
  output: {
    path: path.resolve(__dirname, 'docs'),
    filename: 'bundle.js',
  },
  mode: 'production',
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: { loader: 'babel-loader' },
      },
      {
        test: /\.html$/,
        use: 'html-loader',
      },
      {
        test: /\.(png|jpe?g|gif|ico|svg)$/i,
        type: 'asset/resource',
        generator: {
          filename: 'assets/[name][ext]',
        },
      },
    ],
  },
};
