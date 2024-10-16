// 导入插件
import { nodeResolve } from '@rollup/plugin-node-resolve' // 解析node_modules中的模块
import commonjs from '@rollup/plugin-commonjs' // 转换CommonJS模块为ES模块
import { babel } from '@rollup/plugin-babel' // 使用Babel转换代码
import eslint from '@rollup/plugin-eslint' // 使用ESLint检查代码
import terser from '@rollup/plugin-terser' // 使用Terser压缩代码
import globals from 'rollup-plugin-node-globals' // 提供全局变量
import builtins from 'rollup-plugin-node-builtins' // 提供内置模块

// 定义警告处理函数
const onwarn = warning => {
  // 如果是循环依赖警告，则不处理
  if (warning.code === 'CIRCULAR_DEPENDENCY') return

  // 打印警告信息
  console.warn(`(!) ${warning.message}`) // eslint-disable-line
}

// 导出配置
export default {
  // 输入文件
  input: 'src/pptxtojson.js',
  // 警告处理函数
  onwarn,
  // 输出配置
  output: [
    {
      // 输出文件
      file: 'dist/index.umd.js',
      // 输出格式
      format: 'umd',
      // 输出模块名
      name: 'pptxtojson',
      // 是否生成sourcemap
      sourcemap: true,
    },
    {
      // 输出文件
      file: 'dist/index.js',
      // 输出格式
      format: 'es',
      // 是否生成sourcemap
      sourcemap: true,
    },
  ],
  // 插件配置
  plugins: [
    // 解析node_modules中的模块
    nodeResolve({
      preferBuiltins: false,
    }),
    // 转换CommonJS模块为ES模块
    commonjs(),
    // 使用ESLint检查代码
    eslint(),
    // 使用Babel转换代码
    babel({
      babelHelpers: 'runtime',
      exclude: ['node_modules/**'],
    }),
    // 使用Terser压缩代码
    terser(),
    // 提供全局变量
    globals(),
    // 提供内置模块
    builtins(),
  ]
}