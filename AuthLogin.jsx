import { useState } from "react";
import { supabase } from "./supabaseClient";

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
        <div style={{ fontSize: 20, fontWeight: 800, marginBottom: 4, color: INK, textAlign: "center" }}>
          Kırlı Oto Kurtarma
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
