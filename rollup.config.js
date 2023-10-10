const fs = require('fs');
const { babel } = require('@rollup/plugin-babel');
const { nodeResolve } = require('@rollup/plugin-node-resolve');
const { default: typescript } = require('@rollup/plugin-typescript');
const { default: dotenv } = require('rollup-plugin-dotenv');
const { default: resolve } = require('@rollup/plugin-node-resolve');
const { default: commonjs } = require('@rollup/plugin-commonjs');
const { default: terser } = require('@rollup/plugin-terser');
const { default: replace } = require('@rollup/plugin-replace');

const extensions = ['.ts', '.js', '.tsx'];

const preventThreeShakingPlugin = () => {
  return {
    name: 'no-threeshaking',
    resolveId(id, importer) {
      // let's not theeshake entry points, as we're not exporting anything in App Scripts
      if (!importer) return { id, moduleSideEffects: 'no-treeshake' };
    },
  };
};

module.exports = [
  // Google Apps Script
  {
    input: 'src/apps/index.ts',
    output: [
      {
        dir: 'build',
        format: 'cjs',
      },
    ],
    plugins: [
      typescript({ tsconfig: './tsconfig.build.json' }),
      preventThreeShakingPlugin(),
      nodeResolve({
        extensions,
        mainFields: ['jsnext:main', 'main'],
      }),
      babel({
        extensions,
        babelHelpers: 'runtime',
        comments: false,
      }),
      dotenv(),
    ],
  },

  // React App For Config Form
  {
    input: 'src/config/index.tsx',
    output: [
      {
        file: 'app/assets/javascript.html',
        format: 'iife',
        sourcemap: false,
      },
    ],
    plugins: [
      typescript({ tsconfig: './tsconfig.build.json', jsx: 'react' }),
      resolve(),
      commonjs(),
      replace({
        'process.env.NODE_ENV': `"${process.env.NODE_ENV}"`,
        preventAssignment: true,
      }),
      terser(),
      {
        name: 'wrapInScriptTags',
        renderChunk: (code) => ({
          code: `<script>
${code}
</script>`,
        }),
        // if in dev mode, write to dev.html
        writeBundle() {
          if (process.env.NODE_ENV === 'development') {
            const js = fs.readFileSync('app/assets/javascript.html').toString();

            const configHtml = fs
              .readFileSync('app/assets/config.html')
              .toString();

            const parsedHtml = configHtml
              .split(`<?!= include('app/assets/javascript.html'); ?>`)
              .join(js);

            fs.writeFileSync('app/assets/dev.html', parsedHtml);
          }
        },
      },
    ],
  },
];
