import globals from "globals";
import pluginJs from "@eslint/js";
import tseslint from "typescript-eslint";
import pluginReactConfig from "eslint-plugin-react/configs/recommended.js";


export default [
  {
    settings: {
      "react": {
        "version": "detect",
      }
    }
  },
  {files: ["**/*.{js,mjs,cjs,ts,jsx,tsx}"]},
  { languageOptions: { parserOptions: { ecmaFeatures: { jsx: true } } } },
  {languageOptions: { globals: globals.browser }},
  pluginJs.configs.recommended,
  ...tseslint.configs.recommended,
  pluginReactConfig,
  { files: ["**/*.tsx"],
    rules: {
      '@typescript-eslint/ban-types': [
        'error',
        {
          extendDefaults: true,
          types: {
            /*
              Fine to exclude {} for React
              See https://github.com/typescript-eslint/typescript-eslint/issues/2063#issuecomment-675156492
            */
            '{}': false,
          },
        },
      ],
    },
  },
];
