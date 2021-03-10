import babel from '@rollup/plugin-babel'
import typescript from '@rollup/plugin-typescript'
import { nodeResolve } from '@rollup/plugin-node-resolve'

export default {
  input: "ts/worker.ts",
  output:{
  	format: "iife",
  	file: "json2excel.js",
    banner: `importScripts('https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.2.1/exceljs.min.js');`,
    globals: {
      'exceljs': 'ExcelJS',
    },
  },
  external: ['exceljs'],
  plugins: [
    babel({
      babelHelpers: 'bundled'
    }),
    typescript(),
    nodeResolve({
      mainFields: ['module', 'browser', 'main']
    })
  ],
  moduleContext:() => "self"
};