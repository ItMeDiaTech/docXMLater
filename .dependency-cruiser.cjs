/** @type {import('dependency-cruiser').IConfiguration} */
module.exports = {
  forbidden: [
    {
      name: 'no-circular',
      severity: 'error',
      comment: 'Circular dependencies cause initialization issues and tight coupling.',
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
          '(^|/)\\.[^/]+',       // dot files
          '\\.d\\.ts$',           // type declarations
          'types\\.ts$',          // type-only files
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
