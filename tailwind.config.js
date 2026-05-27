/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./app/**/*.{ts,tsx}', './components/**/*.{ts,tsx}', './public/**/*.html'],
  theme: {
    extend: {
      colors: {
        pc: { blue: '#003b7a', dark: '#002752', light: '#e6eef7', red: '#c0392b' },
      },
      fontFamily: { sans: ['Inter', 'system-ui', 'sans-serif'] },
    },
  },
  plugins: [],
};
