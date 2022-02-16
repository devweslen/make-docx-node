module.exports = {
    env: {
        es2021: true,
        node: true,
    },
    extends: [
        'airbnb-base',
        'eslint:recommended',
    ],
    parser: '@typescript-eslint/parser',
    parserOptions: {
        ecmaVersion: 12,
        sourceType: 'module',
    },
    plugins: [
        '@typescript-eslint',
    ],
    rules: {
        'import/extensions': 'off',
        'import/no-unresolved': 'off',
        'import/no-dynamic-require': 'off',
        'import/prefer-default-export': 'off',
        'no-shadow': 'off',
        'global-require': 'off',
        'no-unused-vars': 'off',
        'no-underscore-dangle': 'off',
        'class-methods-use-this': 'off',
        'max-len': 'off',
        'no-console': process.env.NODE_ENV === 'production' ? 'warn' : 'off',
        'no-debugger': process.env.NODE_ENV === 'production' ? 'warn' : 'off',
    },
};
