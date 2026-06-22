import baseConfig, { isomorphicGuard } from '../../eslint.config.mjs';

export default [
  ...baseConfig,
  isomorphicGuard, // shipped source must stay DOM-free (Node/server-runnable)
  {
    files: ['**/*.json'],
    rules: {
      '@nx/dependency-checks': [
        'error',
        {
          ignoredFiles: [
            '{projectRoot}/eslint.config.{js,cjs,mjs,ts,cts,mts}',
            '{projectRoot}/rollup.config.{js,ts,mjs,mts,cjs,cts}',
            '{projectRoot}/vitest.config.{js,ts,mjs,mts}',
          ],
        },
      ],
    },
    languageOptions: {
      parser: await import('jsonc-eslint-parser'),
    },
  },
  {
    ignores: ['**/out-tsc'],
  },
];
