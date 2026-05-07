import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { resolve } from "node:path";

const base = "/commercial-offer/";

export default defineConfig({
  base,
  plugins: [
    {
      name: "redirect-root-to-base",
      configureServer(server) {
        server.middlewares.use((req, res, next) => {
          const url = req.url?.split("?")[0] ?? "";
          if (url === "/" || url === "") {
            res.statusCode = 302;
            res.setHeader("Location", base);
            res.end();
            return;
          }
          const baseNoSlash = base.replace(/\/$/, "");
          if (url === baseNoSlash) {
            res.statusCode = 302;
            res.setHeader("Location", base);
            res.end();
            return;
          }
          next();
        });
      },
    },
    react(),
  ],
  resolve: {
    alias: {
      "@": resolve(__dirname, "./src"),
    },
  },
  server: {
    port: 5173,
    open: base,
    // Без этого fetch на http://localhost:5173/api/... отдаёт HTML SPA → срыв JSON в auth/me и пустой UI
    proxy: {
      "/api": {
        target: "http://127.0.0.1:8000",
        changeOrigin: true,
      },
    },
  },
});
