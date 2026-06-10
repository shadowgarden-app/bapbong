import nx from '@nx/eslint-plugin';

export default [
  ...nx.configs['flat/base'],
  ...nx.configs['flat/typescript'],
  ...nx.configs['flat/javascript'],
  {
    ignores: ['**/dist', '**/out-tsc', '**/vitest.config.*.timestamp*'],
  },
  {
    files: ['**/*.ts', '**/*.tsx', '**/*.js', '**/*.jsx'],
    rules: {
      '@nx/enforce-module-boundaries': [
        'error',
        {
          enforceBuildableLibDependency: true,
          allow: ['^.*/eslint(\\.base)?\\.config\\.[cm]?[jt]s$'],
          // Layering borrowed from [ref] (see the plan doc §5.2).
          // A package may only import packages whose tag is in its allow-list.
          // Key invariants: word-layout/contracts stay pure; painter-canvas
          // may NOT import the layout engine (it only consumes ResolvedLayout).
          depConstraints: [
            { sourceTag: 'scope:pure', onlyDependOnLibsWithTags: ['scope:pure'] },
            { sourceTag: 'scope:model', onlyDependOnLibsWithTags: ['scope:pure', 'scope:model'] },
            { sourceTag: 'scope:measuring', onlyDependOnLibsWithTags: ['scope:pure'] },
            {
              sourceTag: 'scope:engine',
              onlyDependOnLibsWithTags: ['scope:pure', 'scope:model', 'scope:measuring'],
            },
            { sourceTag: 'scope:painter', onlyDependOnLibsWithTags: ['scope:pure'] },
            { sourceTag: 'scope:io', onlyDependOnLibsWithTags: ['scope:pure', 'scope:model'] },
            { sourceTag: 'scope:adapter', onlyDependOnLibsWithTags: ['scope:app'] },
            { sourceTag: 'scope:app', onlyDependOnLibsWithTags: ['*'] },
          ],
        },
      ],
    },
  },
  {
    files: [
      '**/*.ts',
      '**/*.tsx',
      '**/*.cts',
      '**/*.mts',
      '**/*.js',
      '**/*.jsx',
      '**/*.cjs',
      '**/*.mjs',
    ],
    // Override or add rules here
    rules: {},
  },
];
