import baseConfig, { isomorphicGuard } from '../../eslint.config.mjs';

export default [
  // dist-shim is a bundler artifact (esbuild output committed for the desktop
  // stage step) — not source; linting it trips no-var etc.
  { ignores: ['dist-shim/**'] },
  ...baseConfig,
  isomorphicGuard, // the headless façade must stay DOM-free (Node/server-runnable)
  {
    files: ['**/*.json'],
    rules: {
      '@nx/dependency-checks': [
        'error',
        {
          ignoredFiles: [
            '{projectRoot}/eslint.config.{js,cjs,mjs,ts,cts,mts}',
            '{projectRoot}/esbuild.config.{js,ts,mjs,mts}',
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
