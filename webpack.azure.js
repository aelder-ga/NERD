const path = require('path');
const webpack = require('webpack');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
module.exports = {
  mode: 'production',
  devtool: false,
  entry: {
    polyfill: ['core-js/stable', 'regenerator-runtime/runtime'],
    taskpane: './src/taskpane/taskpane.js',
    commands: './src/commands/commands.js',
    auth: './src/auth/auth.js',
    'sso-redirect': './src/auth/sso-redirect.js',
  },
  output: { path: path.resolve(__dirname, 'dist'), clean: true },
  plugins: [
    new webpack.DefinePlugin({ NERD_API_BASE: JSON.stringify('') }),
    new HtmlWebpackPlugin({ filename:'taskpane.html', template:'./src/taskpane/taskpane.html', chunks:['polyfill','taskpane'] }),
    new HtmlWebpackPlugin({ filename:'geek.html', template:'./src/taskpane/intake.html', chunks:['polyfill','taskpane'] }),
    new HtmlWebpackPlugin({ filename:'index.html', template:'./src/taskpane/taskpane.html', chunks:['polyfill','taskpane'] }),
    new HtmlWebpackPlugin({ filename:'sso-redirect.html', template:'./src/auth/sso-redirect.html', chunks:['sso-redirect'] }),
    new HtmlWebpackPlugin({ filename:'auth.html', template:'./src/auth/auth.html', chunks:['auth'] }),
    new HtmlWebpackPlugin({ filename:'commands.html', template:'./src/commands/commands.html', chunks:['polyfill','commands'] }),
    new CopyWebpackPlugin({ patterns:[
      { from:'assets', to:'assets' },
      { from:'sharepoint/geek-launcher/azure', to:'sharepoint/geek-launcher' },
      { from:'sharepoint/dork-dashboard/azure', to:'sharepoint/dork-dashboard' },
      { from:'sharepoint/dork-dashboard/dork-dashboard.sppkg.base64', to:'downloads/dork-dashboard.sppkg', transform: content => Buffer.from(content.toString().trim(), 'base64') },
      { from:'sharepoint/geek-launcher/geek-launcher.sppkg.base64', to:'downloads/geek-launcher.sppkg', transform: content => Buffer.from(content.toString().trim(), 'base64') },
      { from:'src/taskpane/taskpane.css', to:'taskpane.css' },
      { from:'staticwebapp.config.json', to:'staticwebapp.config.json' },
    ] }),
  ],
};
