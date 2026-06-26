import nx from '@nx/eslint-plugin';

/** DOM globals forbidden in the isomorphic layer (contracts / model / docx) so
 *  its shipped source runs on Node/server too. Spread into those packages' own
 *  eslint.config.mjs — nx lints each project with the project root as the base
 *  path, so this lives there (project-relative globs), not in the root config.
 *  Tests are exempt (they may use jsdom). */
export const isomorphicGuard = {
  files: ['**/*.ts'],
  ignores: ['**/*.spec.ts', '**/*.test.ts'],
  rules: {
    'no-restricted-globals': [
      'error',
      ...[
        'document',
        'window',
        'navigator',
        'getComputedStyle',
        'requestAnimationFrame',
        'cancelAnimationFrame',
        'DOMParser',
        'XMLSerializer',
        'Image',
      ].map((name) => ({
        name,
        message: 'Lớp shared phải isomorphic (chạy được trên Node/server) — không chạm DOM.',
      })),
    ],
  },
};

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
            {
              sourceTag: 'scope:pure',
              onlyDependOnLibsWithTags: ['scope:pure'],
            },
            {
              sourceTag: 'scope:model',
              onlyDependOnLibsWithTags: ['scope:pure', 'scope:model'],
            },
            {
              sourceTag: 'scope:measuring',
              onlyDependOnLibsWithTags: ['scope:pure'],
            },
            {
              sourceTag: 'scope:engine',
              onlyDependOnLibsWithTags: [
                'scope:pure',
                'scope:model',
                'scope:measuring',
              ],
            },
            {
              sourceTag: 'scope:painter',
              onlyDependOnLibsWithTags: ['scope:pure'],
            },
            {
              sourceTag: 'scope:selection',
              onlyDependOnLibsWithTags: ['scope:pure'],
            },
            {
              sourceTag: 'scope:input',
              onlyDependOnLibsWithTags: ['scope:pure', 'scope:model'],
            },
            {
              sourceTag: 'scope:io',
              onlyDependOnLibsWithTags: ['scope:pure', 'scope:model'],
            },
            {
              // Headless meta-tier: aggregates the isomorphic packages
              // (contracts/model/docx/commands) into one backend façade. Stays
              // DOM-free — must NOT pull editor/view (which drag in canvas/DOM).
              sourceTag: 'scope:headless',
              onlyDependOnLibsWithTags: [
                'scope:pure',
                'scope:model',
                'scope:io',
                'scope:headless',
              ],
            },
            {
              // Render tier: load → layout → paint → scroll/zoom/virtualize +
              // geometry (RenderCore) and the read-only viewer. Touches the DOM
              // (canvas) but NOT the input-bridge (scope:input) — so the preview
              // bundle stays free of the ProseMirror editing surface. May pull
              // the a11y mirror (so both viewer + editor are accessible).
              sourceTag: 'scope:view',
              onlyDependOnLibsWithTags: [
                'scope:pure',
                'scope:model',
                'scope:io',
                'scope:measuring',
                'scope:engine',
                'scope:painter',
                'scope:selection',
                'scope:a11y',
                'scope:view',
              ],
            },
            {
              // Accessibility: a visually-hidden ARIA DOM mirror of the document
              // (DOMSerializer over the schema). Pure DOM + the doc model — no
              // canvas/layout/editing deps.
              sourceTag: 'scope:a11y',
              onlyDependOnLibsWithTags: ['scope:pure', 'scope:a11y'],
            },
            {
              sourceTag: 'scope:plugin',
              onlyDependOnLibsWithTags: ['scope:pure'],
            },
            {
              sourceTag: 'scope:adapter',
              onlyDependOnLibsWithTags: ['scope:app'],
            },
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
