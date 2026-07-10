const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    },
    rules: {
      // Known false-positive generator on sequential awaited state machines
      // (https://github.com/eslint/eslint/issues/11899). The workflow engine
      // intentionally mutates step state across awaits; runs are serialized
      // per job by WorkflowEngine._runningItems.
      'require-atomic-updates': 'off',
      // The UPC_* list contracts (StepsJson, PayloadJson, ScheduledFor, ...)
      // round-trip JSON where absent-vs-null is meaningful and SharePoint
      // REST returns null; models mirror that deliberately.
      '@rushstack/no-new-null': 'off',
      // `void somePromise` as a statement is the intentional fire-and-forget
      // marker (pairs with @typescript-eslint/no-floating-promises).
      'no-void': ['warn', { allowAsStatement: true }]
    }
  }
];
