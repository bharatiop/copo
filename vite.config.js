import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANT: change "your-repo-name" below to your actual GitHub repository name.
// If this will be your GitHub *user/org* homepage repo (named yourname.github.io),
// set base to "/" instead.
export default defineConfig({
  plugins: [react()],
  base: "/copo/",
});
