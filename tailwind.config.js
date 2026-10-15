/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        workspace: {
          dark: '#0f141c',
          sidebar: '#141a24',
          panel: '#192231',
          border: '#243044',
          hover: '#222d3e',
          active: '#2c3a50',
          accent: '#3b82f6',
          accentHover: '#2563eb',
          text: '#f1f5f9',
          muted: '#94a3b8'
        }
      }
    },
  },
  plugins: [],
}
