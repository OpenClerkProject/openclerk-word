/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
// Defaults to the project's GitHub Pages deployment. To self-host instead (e.g. for an
// internal/IT-managed deployment), set OPENCLERK_HOST_URL to your own HTTPS URL before
// running `npm run build`/`npm run package` -- see README.md > "Self-hosting".
const urlProd = (process.env.OPENCLERK_HOST_URL || "https://openclerkproject.github.io/openclerk-word/").replace(/\/*$/, "/");

// Couples the packaged manifest's <Version> to the GitHub release tag that triggered the build.
// CI's publish job (.github/workflows/ci.yml) derives this as "<tag-without-v>.0" and sets it
// before running `npm run build`; local/dev/PR builds leave it unset and keep whatever <Version>
// is checked into manifest.xml. Validated here (not just in CI) because this file can be invoked
// directly by anyone running `npm run build`/`npm run package` with the env var set by hand.
const manifestVersion = process.env.OPENCLERK_MANIFEST_VERSION;
if (manifestVersion && !/^\d+\.\d+\.\d+\.0$/.test(manifestVersion)) {
  throw new Error(
    `OPENCLERK_MANIFEST_VERSION must look like "X.Y.Z.0" (trailing segment always 0), got: "${manifestVersion}"`
  );
}

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
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
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
            // README.md links to these with a plain relative path (store-assets/screenshots/...)
            // so the same markdown renders correctly both on GitHub.com and, via build-docs.js, on
            // the deployed GitHub Pages site -- which requires the files to exist at that same
            // relative path under dist/.
            from: "store-assets/screenshots/*.png",
            to: "store-assets/screenshots/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              let result = content.toString();
              if (!dev) {
                result = result.replace(new RegExp(urlDev, "g"), urlProd);
              }
              if (manifestVersion) {
                result = result.replace(/<Version>[^<]*<\/Version>/, `<Version>${manifestVersion}</Version>`);
              }
              return result;
            },
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