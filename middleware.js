import { authMiddleware } from "@clerk/nextjs";

export default authMiddleware({
  publicRoutes: [
    "/",
    "/pricing",
    "/sign-in",
    "/sign-up",
    "/api/stripe/webhook",
    "/api/trial/check",
    "/api/trial/use",
    "/api/tailor/trial",
    "/api/tailor-docx", // made public — route now does its own auth/trial check internally (see tailor-docx.js)
  ],
});

export const config = {
  matcher: ["/((?!.*\\..*|_next).*)", "/", "/(api|trpc)(.*)"],
};
