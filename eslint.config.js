import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import globals from 'globals';

export default tseslint.config(
  {
    ignores: ['dist/**', 'coverage/**', 'node_modules/**', 'out/**', '*.config.js', '*.config.mjs'],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: {
        ...globals.node,
      },
    },
    rules: {
      '@typescript-eslint/no-unused-vars': [
        'error',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
      ],
      '@typescript-eslint/no-explicit-any': 'warn',
      'no-console': ['warn', { allow: ['warn', 'error'] }],
      // `== null` / `!= null` is the intentional null-and-undefined idiom used
      // throughout the oxml layer; require `===` everywhere else.
      eqeqeq: ['error', 'always', { null: 'ignore' }],
      'prefer-const': 'error',
      // OOXML round-trip tests carry literal BOM / preserve-whitespace chars in
      // template/string fixtures on purpose.
      'no-irregular-whitespace': ['error', { skipStrings: true, skipTemplates: true }],
    },
  },
  {
    files: ['**/*.test.ts', '**/*.spec.ts', 'tests/**/*.ts'],
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
      'no-console': 'off',
      // tests intentionally "touch" getters as bare expression statements to
      // assert reads don't mutate — that's the pattern under test.
      '@typescript-eslint/no-unused-expressions': 'off',
    },
  },
  {
    files: ['scripts/**/*.mjs', 'scripts/**/*.js'],
    rules: {
      'no-console': 'off',
    },
  },
);
