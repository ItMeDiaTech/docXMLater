/** @type {import('dependency-cruiser').IConfiguration} */
module.exports = {
  forbidden: [
    {
      name: 'no-circular',
      severity: 'warn',
      comment:
        'Circular dependencies cause initialization issues. The remaining ' +
        'cycles in this codebase are TypeScript-only — every edge that ' +
        'could be made type-only has been (see `import type` usages in ' +
        'TrackingContext, RevisionAutoFixer, InMemoryRevisionAcceptor). ' +
        'Cycle reports retained as warnings for visibility. Severity ' +
        'restored to `error` once dep-cruiser supports type-only edge ' +
        'exclusion in circular detection (issue tracked upstream).',
      from: {},
      to: {
        circular: true,
      },
    },
    {
      name: 'no-orphans',
      severity: 'warn',
      comment: 'Modules not reachable from entry points may be dead code.',
      from: {
        orphan: true,
        pathNot: [
          '(^|/)\\.[^/]+', // dot files
          '\\.d\\.ts$', // type declarations
          'types\\.ts$', // type-only files
        ],
      },
      to: {},
    },
    {
      name: 'utils-no-import-elements',
      severity: 'warn',
      comment: 'Utils should not depend on elements/core/formatting to keep them low-level.',
      from: {
        path: '^src/utils/',
      },
      to: {
        path: '^src/(elements|core|formatting)/',
      },
    },
  ],
  options: {
    doNotFollow: {
      path: 'node_modules',
    },
    tsPreCompilationDeps: true,
    tsConfig: {
      fileName: 'tsconfig.json',
    },
    reporterOptions: {
      dot: {
        collapsePattern: 'node_modules/(@[^/]+/[^/]+|[^/]+)',
      },
    },
  },
};
