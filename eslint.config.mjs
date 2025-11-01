import js from "@eslint/js";
import globals from "globals";
import tseslint from "typescript-eslint";
import { defineConfig } from "eslint/config";

export default defineConfig([
  {
    files: ["**/*.{js,cjs,ts,cts}"],
    ignores: ["dist/**", "node_modules/**"],
    languageOptions: {
      sourceType: "commonjs",
      globals: globals.node, 
    },
    plugins: {
      "@typescript-eslint": tseslint.plugin,
    },
    extends: [
      js.configs.recommended,
      ...tseslint.configs.recommended,
    ],
    rules: {
      semi: ["error", "always"],
      quotes: ["error", "single"],
      "@typescript-eslint/no-explicit-any": "off",
    },
  },
]);
