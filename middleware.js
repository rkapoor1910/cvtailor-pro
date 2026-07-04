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
  ],
});

export const config = {
  matcher: ["/((?!.*\\..*|_next).*)", "/", "/(api|trpc)(.*)"],
};
