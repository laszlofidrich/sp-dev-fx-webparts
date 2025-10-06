// Tell ESLint how to resolve plugins in SPFx projects
require('@rushstack/eslint-config/patch/modern-module-resolution');

module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
  parserOptions: { tsconfigRootDir: __dirname },
  rules: {
    // Stop failing the CI build on unused vars in this sample
    '@typescript-eslint/no-unused-vars': 'off'
  }
};
