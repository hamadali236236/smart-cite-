const path = require("path");

const HtmlWebpackPlugin = require("html-webpack-plugin");

const CopyWebpackPlugin = require("copy-webpack-plugin");

const devCerts = require("office-addin-dev-certs");


module.exports = async (env, argv) => {

  const isProduction =
    argv.mode === "production";


  const httpsOptions =
    isProduction
      ? undefined
      : await devCerts.getHttpsServerOptions();


  return {

    entry: {

      taskpane: "./src/taskpane.js"

    },


    output: {

      path:
        path.resolve(
          __dirname,
          "dist"
        ),

      filename:
        "[name].js",

      clean: true

    },


    resolve: {

      extensions: [
        ".js"
      ]

    },


    plugins: [

      new HtmlWebpackPlugin({

        template:
          "./src/taskpane.html",

        filename:
          "taskpane.html",

        chunks: [
          "taskpane"
        ]

      }),


      new CopyWebpackPlugin({

        patterns: [

          {
            from: "assets",
            to: "assets",
            noErrorOnMissing: true
          }

        ]

      })

    ],


    devServer: {

      static:
        path.resolve(
          __dirname,
          "dist"
        ),


      server: {

        type: "https",

        options:
          httpsOptions

      },


      port: 3000,

      hot: true,


      headers: {

        "Access-Control-Allow-Origin": "*"

      }

    },


    devtool:
      isProduction
        ? false
        : "source-map",


    mode:
      isProduction
        ? "production"
        : "development"

  };

};
