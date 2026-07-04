/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  experimental: {
    serverComponentsExternalPackages: ["mammoth", "adm-zip"],
  },
}
module.exports = nextConfig
