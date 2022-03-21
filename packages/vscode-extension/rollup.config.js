// rollup.config.js
import { defineConfig } from "rollup";
import json from "@rollup/plugin-json";
import commonjs from "@rollup/plugin-commonjs";
import { nodeResolve } from "@rollup/plugin-node-resolve";
import typescriptPlugin from "rollup-plugin-typescript2";
import typescript from "typescript";
import { terser } from "rollup-plugin-terser";

const es2017Plugins = [
  typescriptPlugin({
    typescript,
    abortOnError: false,
  }),
  nodeResolve(),
  commonjs(),
  json(),
  terser(),
];

export default defineConfig({
  input: "src/extension.ts",
  output: {
    file: "dist/extension.js",
    format: "commonjs",
    sourcemap: false,
  },
  plugins: [...es2017Plugins],
  external: ["vscode"],
  treeshake: {
    moduleSideEffects: false,
  },
});
