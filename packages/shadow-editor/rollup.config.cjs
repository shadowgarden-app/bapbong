const { withNx } = require('@nx/rollup/with-nx');

module.exports = withNx(
  {
    main: './src/index.ts',
    outputPath: './dist',
    tsConfig: './tsconfig.lib.json',
    compiler: 'swc',
    // ESM-only for now. Proper dual ESM+CJS needs distinct .mjs/.cjs
    // extensions (Node treats .js as ESM under "type":"module"); @nx/rollup
    // names CJS "index.cjs.js" which Node would parse as ESM. Revisit when a
    // CommonJS consumer actually needs it. See the plan doc §9.
    format: ['esm'],
  },
  {
    // Provide additional rollup configuration here. See: https://rollupjs.org/configuration-options
    output: { sourcemap: true },
  },
);
