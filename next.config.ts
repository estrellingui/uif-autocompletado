import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Allow large file uploads (PDFs + Excel)
  experimental: {
    serverActions: {
      bodySizeLimit: "20mb",
    },
  },
};

export default nextConfig;
