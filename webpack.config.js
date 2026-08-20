/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

// No trailing slash — manifest.xml uses this origin both bare (SupportUrl,
// AppDomain, GetStarted.LearnMoreUrl) and with a path suffix
// (SourceLocation, icon URLs). A bare-origin regex match correctly handles
// both: it matches the exact bare-URL elements, AND matches as a prefix
// inside the path-suffixed ones, leaving the suffix untouched. A version of
// this constant WITH a trailing slash (as it originally shipped) only
// matched the path-suffixed elements — silently leaving AppDomain and the
// other bare-origin elements pointed at localhost in every production build.
const urlDev = "https://localhost:3000";
const urlProd = "https://mentor-gst.vercel.app"; // stable Vercel production alias (project: mentor-gst) — stays fixed across future redeploys

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.js", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          // Marketing landing page — fully self-contained (inline CSS/JS, no
          // relative asset paths, absolute Formspree/Google Fonts URLs only),
          // so a plain copy is all it needs. Served as "index.html" so it
          // shows at the bare domain root — the add-in itself lives at its
          // own separate paths (taskpane.html, commands.html), so this never
          // conflicts with or affects the sideloaded add-in in any way.
          {
            from: "src/mentor_landing.html/mentor_landing.html",
            to: "index.html",
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
