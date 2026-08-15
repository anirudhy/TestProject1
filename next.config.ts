import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Next's built-in build-time ESLint pass doesn't get along with this
  // project's flat eslint.config.mjs in this Next version (fails with an
  // "Invalid Options" error, unrelated to any real lint issue). Actual
  // linting runs via `pnpm lint`, which uses the flat config directly and
  // passes cleanly — see README.
  eslint: {
    ignoreDuringBuilds: true,
  },
  images: {
    remotePatterns: [
      { protocol: "https", hostname: "**.supabase.co" },
      { protocol: "https", hostname: "logo.clearbit.com" },
    ],
  },
};

export default nextConfig;
