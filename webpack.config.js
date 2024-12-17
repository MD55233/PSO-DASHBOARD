module: {
    rules: [
      {
        test: /\.js$/,
        enforce: "pre",
        use: ["source-map-loader"],
        exclude: [/node_modules\/stylis-plugin-rtl/], // Exclude problematic package
      },
    ],
  },
  