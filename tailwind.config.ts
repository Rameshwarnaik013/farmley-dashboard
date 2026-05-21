import type { Config } from 'tailwindcss';
const config: Config = {
  content: ['./src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        brand: { 50: '#e8eaf6', 100: '#c5cae9', 500: '#3f51b5', 700: '#303f9f', 900: '#1a237e' },
        ach: { darkred: '#c00000', red: '#ff0000', orange: '#ffc000', lightgreen: '#92d050', darkgreen: '#00b050', grey: '#d9d9d9' },
      },
    },
  },
  plugins: [],
};
export default config;
