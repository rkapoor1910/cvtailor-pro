import { SignIn } from "@clerk/nextjs";
export default function SignInPage() {
  return (
    <div style={{ minHeight:"100vh", background:"linear-gradient(135deg,#0f172a,#1e1b4b)", display:"flex", alignItems:"center", justifyContent:"center" }}>
      <SignIn path="/sign-in" routing="path" signUpUrl="/sign-up" />
    </div>
  );
}
