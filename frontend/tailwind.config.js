/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        brand: {
          50: "#e8f0ff",
          100: "#d7e4ff",
          500: "#2563eb",
          700: "#1d4ed8"
        }
      }
    }
  },
  plugins: []
};
