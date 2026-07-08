import type { Config } from "tailwindcss";

const config: Config = {
  content: ["./app/**/*.{js,ts,jsx,tsx,mdx}", "./components/**/*.{js,ts,jsx,tsx,mdx}"],
  theme: {
    extend: {
      boxShadow: {
        soft: "0 24px 70px rgba(15, 23, 42, 0.10)",
        panel: "0 18px 50px rgba(16, 24, 40, 0.08)",
        float: "0 16px 38px rgba(16, 24, 40, 0.12)",
        blue: "0 14px 28px rgba(21, 89, 232, 0.20)"
      }
    }
  },
  plugins: []
};

export default config;
