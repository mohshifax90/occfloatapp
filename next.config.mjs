/** @type {import('next').NextConfig} */
const buildId =
  process.env.VERCEL_GIT_COMMIT_SHA?.slice(0, 8) ||
  process.env.GITHUB_SHA?.slice(0, 8) ||
  new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14)

const nextConfig = {
  env: {
    NEXT_PUBLIC_BUILD_ID: buildId,
  },
}

export default nextConfig
