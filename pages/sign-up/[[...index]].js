import { SignUp } from "@clerk/nextjs";
export default function SignUpPage() {
  return (
    <div style={{ minHeight:"100vh", background:"linear-gradient(135deg,#0f172a,#1e1b4b)", display:"flex", alignItems:"center", justifyContent:"center" }}>
      <SignUp path="/sign-up" routing="path" signInUrl="/sign-in" />
    </div>
  );
}
