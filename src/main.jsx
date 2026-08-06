import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";

// Polyfill de `window.storage`: o componente foi originalmente escrito para
// rodar num ambiente que expõe essa API. Aqui reimplementamos a mesma
// interface (get/set assíncronos) usando localStorage do navegador, para que
// o app funcione normalmente em qualquer hospedagem (Vercel, GitHub Pages, etc).
if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    async get(key) {
      try {
        const value = window.localStorage.getItem(key);
        return { value };
      } catch {
        return { value: null };
      }
    },
    async set(key, value) {
      try {
        window.localStorage.setItem(key, value);
        return true;
      } catch {
        return false;
      }
    },
  };
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
