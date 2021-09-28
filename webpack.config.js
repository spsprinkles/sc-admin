var path = require("path");

// Export the configuration
module.exports = (env, argv) => {
    var isDev = argv.mode === "development";

    // Return the configuration
    return {
        // Main project file
        entry: [
            "./node_modules/gd-sprest-bs/dist/gd-sprest-bs.js",
            "./src/index.ts"
        ],

        // Ignore the gd-sprest libraries from the bundle
        externals: {
            "gd-sprest": "$REST",
            "gd-sprest-bs": "$REST"
        },

        // Output information
        output: {
            path: path.resolve(__dirname, "dist"),
            filename: "sc-admin" + (isDev ? "" : ".min") + ".js"
        },

        // Resolve the file names
        resolve: {
            extensions: [".js", ".ts"]
        },

        // Compiler Information
        module: {
            rules: [
                // Handle SASS Files
                {
                    test: /\.css$/,
                    use: [
                        // Create the style nodes from the CommonJS code
                        { loader: "style-loader" },
                        // Translate css to CommonJS
                        { loader: "css-loader" }
                    ]
                },
                // Handle TypeScript Files
                {
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: [
                        // Step 2 - Compile JavaScript ES6 to JavaScript Current Standards
                        {
                            loader: "babel-loader",
                            options: {
                                presets: ["@babel/preset-env"]
                            }
                        },
                        // Step 1 - Compile TypeScript to JavaScript ES6
                        {
                            loader: "ts-loader"
                        }
                    ]
                }
            ]
        }
    };
}