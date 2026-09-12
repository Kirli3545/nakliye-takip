import { useState } from "react";
import { supabase } from "./supabaseClient";

const KIRLI_WORDMARK_SRC = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPUAAAArCAYAAACghc46AAAEIUlEQVR4nO2dzXHbMBBGn2I3oJQgn3NSSlBKcEqgSnBKsEqQSohKiEuILr5HLbCBTJIDiKHiOGMKBD6AxL4ZjT0ekwQX+rD4WSwWz4ubDWlpgVOC+4aW+2V51O+f+nkxSVF3U/2+xSL5+y+eFze/Ez/kCfgU8X5L4BFoAq5tu7JcVrr6/VM/LxUn4Nz9POHeK4Spfd9ik/z9bxM/IDYr4CuwDrj2hKvsNmqJ6mHdfe4v/nYAjoQL3EjAu9wFuIIN8B0TdEk0wLfuM6VhxayZiqj9l2cZcO0RE3RqNrj6echdEGMaot53nxAOwGdM0CoeCa8rIxIli3qJ626HTIiBE/Q2XnGMgTSE15kRgVJFvQZ+EDZ+BidmE3Q+9oTXnTGSEme/G8Z14bY4L10rJ+BLwHVehJvu95D5i0v2wMeR9zACKE3Ue8K7bi1O0Md4xZkkLWFLTP6aXffzHlcXobPa6+76mhvYLJTS/R47fvZBJbULOiZ+1SDE63tsbJ2BEkQ9dvz8WpSYEY8d4cL2ASuGkNyibnAeOnT8dgLuMEGnZocLEQ3BglLE5BT1mPVn6D20rUFrMFFPhByiXuKij8aOt5ZY124KmKjFqEW9xnW3Y1W0hSUaxguUovbx26uI99xgnkDBmF6RzXeIUYnaxwSPDWh4DfPWaVnjtruG1p3NeYhRBJ+k9qb+/ran17GhrEQMVi9ici9pxeL+7X8xMmGiFjMXUTfEHasbcfDpjwwhcxE12Ni6RMaEmBqBzEnU5q3LwucvM8TMSdRg3roEWpyHtv3smVDMfvs8zCvSe9KGcXHKc0Bp78sc2z6F8BFbxsqKQtQ+k6ePJkvNA3V7CbW9LQdcYSi73yc0G+ZtbO1Q2HuJ7ZkuDvWYeoemVbextUNh7wfSRAoagahFfaZPl5MS89YOhb39MUhDSd3IVF/vOWa/D5i3VqKw9zWNaOpgFMUEYdHkELVf8kiNeWuHyt5DG1FFhFnVDXqudeoDmmWnqiv3AoW9G4Ztz1SIuuoDBXIGnyiWncxb9yjsPWRsrVrH3lNpo55T1E9odvBUWbGvoLD30G22islScI2MTz1dzQx97kPnVQESd/y/+zm3Q+dz23vIoe9LXFpotdDO9DvHVAEzLX/HC8z+0HkfIJF6/FN7lJlHYe8hSStanLe+ZiksBn5mXJkC6wnxKSUlbOhQBEhU1f16A4W9h6R+3mFH8iShBFErA1IMjb1XDLN37YcZJqEEUYMmQMLCGXtU9h7CFreObptCIlGKqBUBErb5oEdh7xXDhb3DHXtrXjsCpYgaNAES5q17SrP3Gee13+MaHEtYGMgt6Y13TQTRFpfwPxXeW1+OKdXvX7u938LPjPtrNvSRalM8uEFe/4vnxU3iZ1xN6oq7zNZhmL0nzYdfP//52x+sVvFToUFpDQAAAABJRU5ErkJggg==";

const BRAND = "#EE1820";
const BRAND_DARK = "#A6151B";
const INK = "#1C1B19";
const BG = "#F3F2EE";
const CARD = "#FFFFFF";
const BORDER = "#E1DED5";
const RED = "#B23A2E";
const MUTED = "#6B685F";

const inputStyle = {
  width: "100%",
  boxSizing: "border-box",
  padding: "10px 12px",
  borderRadius: 8,
  border: `1px solid ${BORDER}`,
  fontSize: 14,
  fontFamily: "inherit",
  background: "#FAF9F6",
};

export default function AuthLogin() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function handleLogin(e) {
    e.preventDefault();
    setError("");
    if (!email.trim() || !password) {
      setError("E-posta ve şifre girin.");
      return;
    }
    setLoading(true);
    const { error: authError } = await supabase.auth.signInWithPassword({
      email: email.trim(),
      password,
    });
    setLoading(false);
    if (authError) {
      setError("Giriş başarısız: e-posta veya şifre hatalı.");
    }
    // Başarılıysa App.jsx'teki onAuthStateChange dinleyicisi devreye girer,
    // burada ekstra bir şey yapmaya gerek yok.
  }

  return (
    <div
      style={{
        fontFamily: "'Inter', sans-serif",
        background: BG,
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 24,
      }}
    >
      <form
        onSubmit={handleLogin}
        style={{
          width: "100%",
          maxWidth: 380,
          background: CARD,
          border: `1px solid ${BORDER}`,
          borderRadius: 12,
          padding: 24,
        }}
      >
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 16 }}>
          <img src={KIRLI_WORDMARK_SRC} alt="Kırlı" style={{ height: 32, width: "auto" }} />
        </div>
        <div style={{ fontSize: 13, color: MUTED, marginBottom: 20, textAlign: "center" }}>
          Nakliye ve Görev Yönetim Sistemi
        </div>

        <label style={{ display: "block", marginBottom: 14 }}>
          <span style={{ display: "block", fontSize: 13, color: MUTED, marginBottom: 5, fontWeight: 600 }}>
            E-posta
          </span>
          <input
            type="email"
            autoComplete="username"
            style={inputStyle}
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="ornek@kirliotokurtarma.com"
          />
        </label>

        <label style={{ display: "block", marginBottom: 14 }}>
          <span style={{ display: "block", fontSize: 13, color: MUTED, marginBottom: 5, fontWeight: 600 }}>
            Şifre
          </span>
          <input
            type="password"
            autoComplete="current-password"
            style={inputStyle}
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="••••••••"
          />
        </label>

        {error && <div style={{ color: RED, fontSize: 13, marginBottom: 12 }}>{error}</div>}

        <button
          type="submit"
          disabled={loading}
          style={{
            width: "100%",
            border: "none",
            cursor: loading ? "default" : "pointer",
            fontWeight: 700,
            borderRadius: 8,
            padding: "11px 16px",
            fontSize: 14,
            background: loading ? "#D8D5CC" : BRAND,
            color: "#fff",
          }}
        >
          {loading ? "Giriş yapılıyor..." : "Giriş Yap"}
        </button>

        <div style={{ fontSize: 12, color: MUTED, marginTop: 16, textAlign: "center" }}>
          Hesabınız yoksa yöneticinizden hesap açmasını isteyin.
        </div>
      </form>
    </div>
  );
}
