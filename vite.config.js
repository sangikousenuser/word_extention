import { defineConfig } from "vite";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [basicSsl()],
  base: process.env.VITE_BASE_PATH || "/",
  server: {
    https: true,
    host: "localhost",
    port: 3000,
  },
  build: {
    rollupOptions: {
      input: {
        commands: "commands.html",
        taskpane: "taskpane.html",
      },
    },
  },
});
