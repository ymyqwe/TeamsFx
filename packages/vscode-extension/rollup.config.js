// rollup.config.js
import { defineConfig } from "rollup";
import json from "@rollup/plugin-json";
import commonjs from "@rollup/plugin-commonjs";
import { nodeResolve } from "@rollup/plugin-node-resolve";
import typescriptPlugin from "rollup-plugin-typescript2";
import typescript from "typescript";
import { terser } from "rollup-plugin-terser";
import { visualizer } from "rollup-plugin-visualizer";

const es2017Plugins = [
  typescriptPlugin({
    typescript,
    abortOnError: false,
  }),
  nodeResolve(),
  commonjs(),
  json(),
  visualizer({
    template: "sunburst",
  }),
];

export default defineConfig({
  input: "src/extension.ts",
  output: {
    file: "dist/extension.js",
    format: "commonjs",
    sourcemap: true,
  },
  plugins: [...es2017Plugins],
  external: ["vscode"],
  treeshake: {
    moduleSideEffects: false,
  },
});
