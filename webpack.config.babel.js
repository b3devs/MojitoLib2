import path from 'path';
import webpack from 'webpack';

export default {
  entry: [path.join(__dirname, 'src/index.js')],
  mode: "development",
  devtool: false,
  module: {
    rules: [{
      exclude: /node_modules/,
      test: /\.js$/,
      use: { loader: 'babel-loader' },
    }],
  },
  output: {
    filename: 'MojitoLib.js',
    library: {
      // name: 'MojitoLib',
      type: 'this',
    },
    path: path.join(__dirname, 'dist'),
    environment: {
      // The environment supports arrow functions ('() => { ... }').
      arrowFunction: false,
      // The environment supports BigInt as literal (123n).
      bigIntLiteral: false,
      // The environment supports const and let for variable declarations.
      const: true,
      // The environment supports destructuring ('{ a, b } = obj').
      destructuring: true,
      // The environment supports an async import() function to import EcmaScript modules.
      dynamicImport: false,
      // The environment supports 'for of' iteration ('for (const x of array) { ... }').
      forOf: true,
      // The environment supports ECMAScript Module syntax to import ECMAScript modules (import ... from '...').
      module: false,
      // The environment supports optional chaining ('obj?.a' or 'obj?.()').
      optionalChaining: true,
      // The environment supports template literals.
      templateLiteral: true,
    },
  },
  resolve: {
    extensions: ['.js'],
    modules: [
      path.join(__dirname, 'src'),
      path.join(__dirname, 'node_modules'),
    ],
  },
};
