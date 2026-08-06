import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: [
        "icons/icon-192.png",
        "icons/icon-512.png",
        "fonts/Inter-Variable.woff2",
        "fonts/Tomorrow-Light.woff2",
        "fonts/Tomorrow-Regular.woff2",
        "fonts/Tomorrow-Medium.woff2",
        "fonts/Tomorrow-SemiBold.woff2",
        "fonts/Tomorrow-Bold.woff2"
      ],
      manifest: {
        name: "Patrimônio Dashboard",
        short_name: "Patrimônio",
        description: "Dashboard pessoal de investimentos e patrimônio",
        theme_color: "#13171B",
        background_color: "#13171B",
        display: "standalone",
        orientation: "portrait",
        start_url: "/",
        scope: "/",
        icons: [
          { src: "/icons/icon-192.png", sizes: "192x192", type: "image/png", purpose: "any" },
          { src: "/icons/icon-512.png", sizes: "512x512", type: "image/png", purpose: "any" },
          { src: "/icons/icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" }
        ]
      },
      workbox: {
        globPatterns: ["**/*.{js,css,html,png,svg,ico,woff2}"]
      }
    })
  ]
});
