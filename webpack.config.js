var GasPlugin = require('gas-webpack-plugin')
var es3ify = require('es3ify-webpack-plugin')

module.exports = {
  entry: ['./src/code.js'],
  output: {
    path: `${__dirname}/dist`,
    filename: 'code.js'
  },
  plugins: [
    new GasPlugin(),
    new es3ify()
  ],
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: [
          {
            loader: 'babel-loader',
          }
        ]
      }
    ]
  }
}
