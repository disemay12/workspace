// Пакет объявлен как CommonJS ("type": "commonjs" в package.json),
// поэтому конфигурация PostCSS должна использовать module.exports.
// Tailwind CSS v4 подключается через @tailwindcss/vite, поэтому
// плагин tailwindcss здесь не дублируется.
module.exports = {
  plugins: {
    autoprefixer: {},
  },
};
