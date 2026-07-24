import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base 必須是你的 GitHub 倉庫名稱，前後都要有斜線
export default defineConfig({
  plugins: [react()],
  base: "/invoice-stats/",
});
