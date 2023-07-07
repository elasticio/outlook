module.exports = {
  extends: 'airbnb-base',
  parserOptions: {
    ecmaVersion: 2020,
  },
  env: {
    node: true,
    mocha: true,
  },
  rules: {
    'no-await-in-loop': 0,
    'no-plusplus': 0,
  },
  overrides: [
    {
      files: [
        '*.test.js',
        '*.spec*',
      ],
      rules: {
        'no-unused-expressions': 'off',
      },
    },
    {
      files: [
        '*',
      ],
      rules: {
        'max-len': ['error', { code: 180 }],
      },
    },
  ],
};
