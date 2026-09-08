import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import AuthLogin from "./AuthLogin";

const LOGO_SRC = "data:image/jpeg;base64,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Iiup4KsMg1X/ALKsP+fG1/79L/hVuiiw02tir/ZVh/z42v8A36X/AAo/sqw/58bX/v0v+FWqKVkPnl3Kv9lWH/Pja/8Afpf8KP7KsP8Anytf+/S/4Vaoosg55dyr/ZVh/wA+Vr/36X/Cj+yrD/nytf8Av0v+FWqKLIOeXcq/2VYf8+Vr/wB+l/wo/sqw/wCfK1/79L/hVqiiyDnl3Kv9lWH/AD42v/fpf8KP7KsP+fK1/wC/S/4Vaoosg55dyr/ZVh/z42v/AH6X/Cj+yrD/AJ8bX/v0v+FWqKLIOeXcq/2VYf8APja/9+l/wo/sqw/58bX/AL9L/hVqiiyDnl3Kv9lWH/Pja/8Afpf8KT+yrD/nxtf+/S/4Vboosg55dyqumWSOrpZ2yspyCIlBB/KrIGKWinYTbe4UUUUCCiiigAxiiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD//Z";

const BRAND = "#EE1820";
const BRAND_DARK = "#A6151B";
const STEEL = "#35566B";
const INK = "#1C1B19";
const BG = "#F3F2EE";
const CARD = "#FFFFFF";
const BORDER = "#E1DED5";
const GREEN = "#4B7A3E";
const RED = "#B23A2E";
const MUTED = "#6B685F";

const STATUS_COLORS = {
  "Bekliyor": { bg: "#FCE3E4", text: BRAND_DARK },
  "Yolda": { bg: "#E4ECF1", text: STEEL },
  "Tamamlandı": { bg: "#E7F0E3", text: GREEN },
};

const LEAVE_STATUS_COLORS = {
  "Beklemede": { bg: "#FCE3E4", text: BRAND_DARK },
  "Onaylandı": { bg: "#E7F0E3", text: GREEN },
  "Reddedildi": { bg: "#F6E2E0", text: RED },
};

const DOC_TYPES = [
  { key: "ruhsat", label: "Ruhsat" },
  { key: "muayene", label: "Muayene Belgesi" },
  { key: "periyodikKontrol", label: "Periyodik Kontrol Belgesi" },
  { key: "trafikSigortasi", label: "Trafik Sigortası" },
  { key: "kasko", label: "Kasko" },
];

const EMPLOYEE_DOC_TYPES = [
  { key: "iseGiris", label: "İşe Giriş" },
  { key: "isg", label: "İSG Belgesi" },
  { key: "saglikRaporu", label: "Sağlık Raporu" },
  { key: "kkdFormu", label: "KKD Formu" },
  { key: "ehliyet", label: "Ehliyet" },
  { key: "cekiciOperatorluk", label: "Çekici Operatörlük Belgesi" },
  { key: "adliSicil", label: "Adli Sicil Kaydı" },
];

const SHIPMENT_TYPES = ["Nakliye", "Yükleme/İndirme"];

// Her veri türü kendi kaydında saklanır — böylece biri belge yüklerken
// biri nakliye bildirirse, birbirlerinin verisini ezmezler.
const STORAGE_KEYS = [
  "employees",
  "shipments",
  "leaveRequests",
  "overtimeReports",
  "machines",
  "workReports",
  "vehicles",
  "employeeDocs",
];

function daysUntil(dateStr) {
  if (!dateStr) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(dateStr);
  target.setHours(0, 0, 0, 0);
  return Math.round((target - today) / 86400000);
}

function docStatus(dateStr) {
  const d = daysUntil(dateStr);
  if (d === null) return { label: "Tarih girilmedi", color: { bg: "#EFEDE6", text: MUTED } };
  if (d < 0) return { label: `${Math.abs(d)} gün önce doldu`, color: { bg: "#F6E2E0", text: RED } };
  if (d <= 10) return { label: `${d} gün kaldı`, color: { bg: "#FCE3E4", text: BRAND_DARK } };
  return { label: "Geçerli", color: { bg: "#E7F0E3", text: GREEN } };
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function dataUrlToBlob(dataUrl) {
  const [header, base64] = dataUrl.split(",");
  const mimeMatch = header.match(/data:(.*?);base64/);
  const mime = mimeMatch ? mimeMatch[1] : "application/octet-stream";
  const binary = atob(base64);
  const array = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) array[i] = binary.charCodeAt(i);
  return new Blob([array], { type: mime });
}

function downloadDataUrl(dataUrl, fileName) {
  try {
    const blob = dataUrlToBlob(dataUrl);
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = fileName || "belge";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 15000);
  } catch (e) {
    // Blob dönüşümü başarısız olursa, en azından yeni sekmede açmayı dene
    window.open(dataUrl, "_blank");
  }
}

function safeKeyPart(str) {
  return String(str).trim().replace(/[\s\/\\"']+/g, "_");
}

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function resizeImage(file, maxWidth = 900) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", 0.72));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

const btnBase = {
  border: "none",
  cursor: "pointer",
  fontWeight: 600,
  borderRadius: 8,
  padding: "10px 16px",
  fontSize: 14,
};

function PrimaryButton({ children, onClick, style, type = "button", disabled }) {
  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled}
      style={{
        ...btnBase,
        background: disabled ? "#D8D5CC" : BRAND,
        color: "#fff",
        opacity: disabled ? 0.7 : 1,
        ...style,
      }}
    >
      {children}
    </button>
  );
}

function GhostButton({ children, onClick, style }) {
  return (
    <button
      type="button"
      onClick={onClick}
      style={{
        ...btnBase,
        background: "transparent",
        color: INK,
        border: `1px solid ${BORDER}`,
        ...style,
      }}
    >
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "block", marginBottom: 14 }}>
      <span style={{ display: "block", fontSize: 13, color: MUTED, marginBottom: 5, fontWeight: 600 }}>
        {label}
      </span>
      {children}
    </label>
  );
}

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

function Badge({ text, color }) {
  return (
    <span
      style={{
        display: "inline-block",
        padding: "3px 10px",
        borderRadius: 999,
        fontSize: 12,
        fontWeight: 700,
        background: color.bg,
        color: color.text,
      }}
    >
      {text}
    </span>
  );
}

function Card({ children, style, onClick }) {
  return (
    <div
      onClick={onClick}
      style={{
        background: CARD,
        border: `1px solid ${BORDER}`,
        borderRadius: 12,
        padding: 16,
        ...style,
      }}
    >
      {children}
    </div>
  );
}

export default function App() {
  const [data, setData] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [session, setSession] = useState(undefined);
  const [authReady, setAuthReady] = useState(false);
  const [tab, setTab] = useState("nakliyeler");
  const [saveError, setSaveError] = useState("");
  const fileInputRef = useRef(null);
  const [activeShipmentForPhoto, setActiveShipmentForPhoto] = useState(null);

  const role = session?.user?.user_metadata?.role || null;
  const currentName = session?.user?.user_metadata?.name || "";

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setAuthReady(true);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    // Oturum durumu netlesmeden ya da giris yapilmadan veri cekmeye CALISMIYORUZ
    if (!session) {
      return;
    }
    setLoaded(false);
    (async () => {
      const base = { employees: [], shipments: [], leaveRequests: [], overtimeReports: [], machines: [], workReports: [], vehicles: [], employeeDocs: {} };
      try {
        const results = await Promise.all(
          STORAGE_KEYS.map((k) => window.storage.get(`data:${k}`, true).catch(() => null))
        );
        let assembled = {};
        let anyFound = false;
        STORAGE_KEYS.forEach((k, i) => {
          if (results[i] && results[i].value !== undefined && results[i].value !== null) {
            try {
              assembled[k] = JSON.parse(results[i].value);
              anyFound = true;
            } catch (e) {}
          }
        });

        if (!anyFound) {
          // Yeni ayrı-anahtar sistemine geçmeden önce eski tekli kayıt var mıydı, ona bak
          // (varsa göç ettir, yoksa boştan başla).
          const legacy = await window.storage.get("workspace-data", true).catch(() => null);
          if (legacy && legacy.value) {
            const legacyData = JSON.parse(legacy.value);
            assembled = { ...base, ...legacyData };
            await Promise.all(
              STORAGE_KEYS.map((k) =>
                window.storage.set(`data:${k}`, JSON.stringify(assembled[k] ?? base[k]), true)
              )
            );
          }
        }

        setData({ ...base, ...assembled });
      } catch (e) {
        setData(base);
      }
      setLoaded(true);
    })();
  }, [session]);

  useEffect(() => {
    if (role === "calisan" && currentName && data && !data.employees.includes(currentName)) {
      persist({ ...data, employees: [...data.employees, currentName] });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [role, currentName, data]);

  async function persist(next) {
    setData(next);
    try {
      const results = await Promise.all(
        STORAGE_KEYS.map((k) => window.storage.set(`data:${k}`, JSON.stringify(next[k]), true))
      );
      if (results.some((r) => !r)) setSaveError("Kaydedilemedi, tekrar deneyin.");
      else setSaveError("");
    } catch (e) {
      setSaveError("Kaydedilemedi, tekrar deneyin.");
    }
  }

  if (!authReady) {
    return <div style={{ padding: 24, color: MUTED, fontFamily: "sans-serif" }}>Yükleniyor...</div>;
  }

  if (!session) {
    return <AuthLogin />;
  }

  if (!loaded) {
    return <div style={{ padding: 24, color: MUTED, fontFamily: "sans-serif" }}>Yükleniyor...</div>;
  }

  if (!role || !currentName) {
    return (
      <div style={{ padding: 24, color: RED, fontFamily: "sans-serif" }}>
        Hesabınıza rol/isim bilgisi tanımlanmamış. Yöneticinizden Supabase panelinde
        hesabınıza "role" ve "name" bilgisi eklemesini isteyin, sonra çıkış yapıp
        tekrar giriş yapın.
        <div style={{ marginTop: 12 }}>
          <button onClick={() => supabase.auth.signOut()} style={{ cursor: "pointer" }}>Çıkış Yap</button>
        </div>
      </div>
    );
  }

  const myShipments = data.shipments.filter((s) => s.assignedTo === currentName);
  const myLeaves = data.leaveRequests.filter((l) => l.employeeName === currentName);

  function updateShipment(id, patch) {
    const next = {
      ...data,
      shipments: data.shipments.map((s) => (s.id === id ? { ...s, ...patch } : s)),
    };
    persist(next);
  }

  function handlePhotoPick(shipmentId) {
    setActiveShipmentForPhoto(shipmentId);
    fileInputRef.current?.click();
  }

  async function onFileChosen(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file || !activeShipmentForPhoto) return;
    const dataUrl = await resizeImage(file, 1600);
    const shipment = data.shipments.find((s) => s.id === activeShipmentForPhoto);
    updateShipment(activeShipmentForPhoto, { photos: [...(shipment.photos || []), dataUrl] });
    setActiveShipmentForPhoto(null);
  }

  return (
    <>
    <div className="no-print" style={{ fontFamily: "'Inter', 'Segoe UI', sans-serif", background: BG, minHeight: "100%", color: INK }}>
      <input type="file" accept="image/*" ref={fileInputRef} onChange={onFileChosen} style={{ display: "none" }} />
      <TopBar
        role={role}
        currentName={currentName}
        onSwitch={() => supabase.auth.signOut()}
      />
      {saveError && (
        <div style={{ background: "#F6E2E0", color: RED, padding: "8px 20px", fontSize: 13 }}>{saveError}</div>
      )}
      {role === "yonetici" && <ExpiryAlertBanner data={data} onGoTo={(tabKey) => setTab(tabKey)} />}
      <div style={{ display: "flex", gap: 6, padding: "12px 20px 0", borderBottom: `1px solid ${BORDER}` }}>
        {[
          ["nakliyeler", role === "yonetici" ? "Nakliyeler" : "Görevlerim"],
          ["izinler", "İzin / Rapor"],
          ["mesai", "Fazla Mesai"],
          ["calismaformu", "Çalışma Formu"],
          ...(role === "yonetici" ? [["calisanlar", "Çalışanlar"], ["araclar", "Araçlar"], ["personelbelgeleri", "Personel Belgeleri"]] : [["belgelerim", "Belgelerim"]]),
        ].map(([key, label]) => (
          <button
            key={key}
            onClick={() => setTab(key)}
            style={{
              border: "none",
              background: "none",
              cursor: "pointer",
              padding: "10px 14px",
              fontSize: 14,
              fontWeight: 600,
              color: tab === key ? BRAND_DARK : MUTED,
              borderBottom: tab === key ? `3px solid ${BRAND}` : "3px solid transparent",
            }}
          >
            {label}
          </button>
        ))}
      </div>

      <div style={{ padding: 20, maxWidth: 760, margin: "0 auto" }}>
        {tab === "nakliyeler" && role === "yonetici" && (
          <ManagerShipments data={data} persist={persist} />
        )}
        {tab === "nakliyeler" && role === "calisan" && (
          <DriverShipments
            shipments={myShipments}
            updateShipment={updateShipment}
            onPhotoPick={handlePhotoPick}
            onCreate={(s) =>
              persist({
                ...data,
                shipments: [
                  { ...s, id: uid(), assignedTo: currentName, source: "surucu", photos: [], createdAt: new Date().toISOString() },
                  ...data.shipments,
                ],
              })
            }
          />
        )}
        {tab === "izinler" && role === "yonetici" && (
          <ManagerLeaves data={data} persist={persist} />
        )}
        {tab === "izinler" && role === "calisan" && (
          <DriverLeaves leaves={myLeaves} onSubmit={(req) => persist({ ...data, leaveRequests: [{ ...req, id: uid(), employeeName: currentName, status: "Beklemede" }, ...data.leaveRequests] })} />
        )}
        {tab === "mesai" && role === "yonetici" && (
          <ManagerOvertime data={data} />
        )}
        {tab === "mesai" && role === "calisan" && (
          <DriverOvertime
            reports={data.overtimeReports.filter((r) => r.employeeName === currentName)}
            onSubmit={(r) => persist({ ...data, overtimeReports: [{ ...r, id: uid(), employeeName: currentName, createdAt: new Date().toISOString() }, ...data.overtimeReports] })}
          />
        )}
        {tab === "calisanlar" && role === "yonetici" && (
          <EmployeeList data={data} persist={persist} />
        )}
        {tab === "araclar" && role === "yonetici" && (
          <VehiclesSection data={data} persist={persist} />
        )}
        {tab === "personelbelgeleri" && role === "yonetici" && (
          <EmployeeDocsSection data={data} persist={persist} />
        )}
        {tab === "belgelerim" && role === "calisan" && (
          <MyDocuments employeeName={currentName} docs={(data.employeeDocs || {})[currentName] || {}} />
        )}
        {tab === "calismaformu" && role === "yonetici" && (
          <ManagerWorkReports data={data} />
        )}
        {tab === "calismaformu" && role === "calisan" && (
          <DriverWorkForm
            machines={data.machines}
            reports={data.workReports.filter((r) => r.employeeName === currentName)}
            onSubmit={(r) => persist({ ...data, workReports: [{ ...r, id: uid(), employeeName: currentName, createdAt: new Date().toISOString() }, ...data.workReports] })}
          />
        )}
      </div>
    </div>

    <style>{`
        @media print {
          .no-print { display: none !important; }
          .print-only { display: block !important; }
        }
        .print-only { display: none; }
      `}</style>
      <div className="print-only" style={{ padding: 24, fontFamily: "'Inter', sans-serif", color: "#000" }}>
        <img src={LOGO_SRC} alt="Mehmet Kırlı" style={{ height: 60, marginBottom: 10 }} />
        <h2 style={{ marginBottom: 4 }}>Fazla Mesai Raporu</h2>
        <div style={{ fontSize: 12, color: "#444", marginBottom: 16 }}>Oluşturulma: {new Date().toLocaleString("tr-TR")}</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr>
              {["Çalışan", "Tarih", "Fazla Mesai (saat)", "Açıklama"].map((h) => (
                <th key={h} style={{ textAlign: "left", borderBottom: "2px solid #000", padding: "6px 8px" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {[...data.overtimeReports].sort((a, b) => (a.date < b.date ? 1 : -1)).map((r) => (
              <tr key={r.id}>
                <td style={{ padding: "6px 8px", borderBottom: "1px solid #ccc" }}>{r.employeeName}</td>
                <td style={{ padding: "6px 8px", borderBottom: "1px solid #ccc" }}>{r.date}</td>
                <td style={{ padding: "6px 8px", borderBottom: "1px solid #ccc" }}>{r.hours}</td>
                <td style={{ padding: "6px 8px", borderBottom: "1px solid #ccc" }}>{r.note}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

function TopBar({ role, currentName, onSwitch }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 20px", borderBottom: `1px solid ${BORDER}` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <img src={LOGO_SRC} alt="Mehmet Kırlı" style={{ height: 40, width: "auto" }} />
        <div>
          <div style={{ fontWeight: 800, fontSize: 15 }}>Nakliye Takip</div>
          <div style={{ fontSize: 12, color: MUTED }}>{currentName} · {role === "yonetici" ? "Yönetici" : "Çalışan"}</div>
        </div>
      </div>
      <GhostButton onClick={onSwitch}>Çıkış</GhostButton>
    </div>
  );
}

function exportShipmentsToExcel(shipments) {
  const rows = shipments.map((s) => ({
    "İşlem Türü": s.title || "",
    "Müşteri": s.musteri || "",
    "Nereden": s.from || "",
    "Nereye": s.to || "",
    "Ekipman": s.equipment || "",
    "Plaka": s.plate || "",
    "Atanan Çalışan": s.assignedTo || "",
    "Tarih": s.date || "",
    "Durum": s.status || "",
    "Bildiren": s.source === "surucu" ? "Şoför" : "Yönetici",
    "Not": s.notes || "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [
    { wch: 16 }, { wch: 22 }, { wch: 18 }, { wch: 18 }, { wch: 20 }, { wch: 14 },
    { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 24 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Nakliyeler");
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `nakliyeler-${dateStr}.xlsx`);
}

function ManagerShipments({ data, persist }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ title: "Nakliye", musteri: "", from: "", to: "", equipment: "", plate: "", assignedTo: data.employees[0] || "", date: "" });
  const [error, setError] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [driverFilter, setDriverFilter] = useState("");
  const [plateFilter, setPlateFilter] = useState("");

  function submit() {
    if (!form.from.trim() || !form.to.trim() || !form.assignedTo) {
      setError("Nereden, nereye ve atanan çalışan alanları zorunlu.");
      return;
    }
    const shipment = {
      id: uid(),
      title: form.title || "Nakliye",
      musteri: form.musteri.trim(),
      from: form.from.trim(),
      to: form.to.trim(),
      equipment: form.equipment.trim(),
      plate: form.plate.trim().toUpperCase(),
      assignedTo: form.assignedTo,
      date: form.date,
      status: "Bekliyor",
      notes: "",
      photos: [],
      createdAt: new Date().toISOString(),
    };
    persist({ ...data, shipments: [shipment, ...data.shipments] });
    setForm({ title: "Nakliye", musteri: "", from: "", to: "", equipment: "", plate: "", assignedTo: data.employees[0] || "", date: "" });
    setError("");
    setShowForm(false);
  }

  const filteredShipments = data.shipments.filter((s) => {
    if (dateFrom && (!s.date || s.date < dateFrom)) return false;
    if (dateTo && (!s.date || s.date > dateTo)) return false;
    if (driverFilter && s.assignedTo !== driverFilter) return false;
    if (plateFilter && !(s.plate || "").toUpperCase().includes(plateFilter.trim().toUpperCase())) return false;
    return true;
  });

  const hasActiveFilter = dateFrom || dateTo || driverFilter || plateFilter;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 18, fontWeight: 700 }}>Nakliyeler</div>
        <PrimaryButton onClick={() => setShowForm((v) => !v)}>{showForm ? "Vazgeç" : "+ Yeni nakliye"}</PrimaryButton>
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: MUTED, marginBottom: 10 }}>FİLTRELE</div>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 140 }}>
            <Field label="Başlangıç">
              <input type="date" style={inputStyle} value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} />
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 140 }}>
            <Field label="Bitiş">
              <input type="date" style={inputStyle} value={dateTo} onChange={(e) => setDateTo(e.target.value)} />
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 140 }}>
            <Field label="Şoför">
              <select style={inputStyle} value={driverFilter} onChange={(e) => setDriverFilter(e.target.value)}>
                <option value="">Tümü</option>
                {data.employees.map((e) => (
                  <option key={e} value={e}>{e}</option>
                ))}
              </select>
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 140 }}>
            <Field label="Plaka">
              <input style={inputStyle} value={plateFilter} onChange={(e) => setPlateFilter(e.target.value)} placeholder="Örn. 35 ABC" />
            </Field>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", marginTop: 4 }}>
          {hasActiveFilter && (
            <GhostButton onClick={() => { setDateFrom(""); setDateTo(""); setDriverFilter(""); setPlateFilter(""); }}>Filtreleri Temizle</GhostButton>
          )}
          <PrimaryButton
            disabled={filteredShipments.length === 0}
            onClick={() => exportShipmentsToExcel(filteredShipments)}
          >
            Excel'e Aktar ({filteredShipments.length})
          </PrimaryButton>
        </div>
      </Card>

      {showForm && (
        <Card style={{ marginBottom: 18 }}>
          {data.employees.length === 0 && (
            <div style={{ fontSize: 13, color: RED, marginBottom: 10 }}>
              Önce Çalışanlar sekmesinden bir çalışan ekleyin.
            </div>
          )}
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="İşlem Türü">
                <select style={inputStyle} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })}>
                  {SHIPMENT_TYPES.map((t) => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Müşteri / Firma">
                <input style={inputStyle} value={form.musteri} onChange={(e) => setForm({ ...form, musteri: e.target.value })} placeholder="Örn. ABC İnşaat Ltd. Şti." />
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Nereden">
                <input style={inputStyle} value={form.from} onChange={(e) => setForm({ ...form, from: e.target.value })} placeholder="Örn. İzmir Şantiye" />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Nereye">
                <input style={inputStyle} value={form.to} onChange={(e) => setForm({ ...form, to: e.target.value })} placeholder="Örn. Manisa OSB" />
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Ekipman / iş makinesi">
                <input style={inputStyle} value={form.equipment} onChange={(e) => setForm({ ...form, equipment: e.target.value })} placeholder="Örn. Lastikli yükleyici" />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Aracın plakası">
                <input style={inputStyle} value={form.plate} onChange={(e) => setForm({ ...form, plate: e.target.value })} placeholder="Örn. 35 ABC 123" />
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Atanan çalışan">
                <select style={inputStyle} value={form.assignedTo} onChange={(e) => setForm({ ...form, assignedTo: e.target.value })}>
                  <option value="">Seçin</option>
                  {data.employees.map((e) => (
                    <option key={e} value={e}>{e}</option>
                  ))}
                </select>
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Tarih">
                <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
              </Field>
            </div>
          </div>
          {error && <div style={{ color: RED, fontSize: 13, marginBottom: 10 }}>{error}</div>}
          <PrimaryButton onClick={submit}>Görevi ata</PrimaryButton>
        </Card>
      )}

      {filteredShipments.length === 0 ? (
        <EmptyState text={data.shipments.length === 0 ? "Henüz nakliye görevi yok. Yeni nakliye ekleyerek başlayın." : "Seçilen filtrelere uyan nakliye yok."} />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {filteredShipments.map((s) => (
            <ShipmentCard key={s.id} s={s} showAssignee />
          ))}
        </div>
      )}
    </div>
  );
}

function ShipmentCard({ s, showAssignee, footer }) {
  const [lightboxPhoto, setLightboxPhoto] = useState(null);
  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15 }}>{s.title}</div>
          {s.musteri && <div style={{ fontSize: 13, color: MUTED, marginTop: 2 }}>Müşteri: {s.musteri}</div>}
          <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>
            {s.from} <span style={{ color: BRAND }}>→</span> {s.to}
          </div>
          {s.equipment && <div style={{ fontSize: 13, color: MUTED, marginTop: 2 }}>Ekipman: {s.equipment}</div>}
          {s.plate && <div style={{ fontSize: 13, color: MUTED, marginTop: 2 }}>Plaka: {s.plate}</div>}
          {showAssignee && <div style={{ fontSize: 13, color: MUTED, marginTop: 2 }}>Atanan: {s.assignedTo}</div>}
          {s.date && <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>Tarih: {s.date}</div>}
          {s.source === "surucu" && (
            <div style={{ marginTop: 6 }}>
              <Badge text="Şoför bildirdi" color={{ bg: "#E4ECF1", text: STEEL }} />
            </div>
          )}
        </div>
        <Badge text={s.status} color={STATUS_COLORS[s.status]} />
      </div>
      {s.photos && s.photos.length > 0 && (
        <div style={{ display: "flex", gap: 8, marginTop: 12, flexWrap: "wrap" }}>
          {s.photos.map((p, i) => (
            <img
              key={i}
              src={p}
              alt="Ürün fotoğrafı"
              onClick={() => setLightboxPhoto(p)}
              style={{ width: 72, height: 72, objectFit: "cover", borderRadius: 8, border: `1px solid ${BORDER}`, cursor: "zoom-in" }}
            />
          ))}
        </div>
      )}
      {footer}
      {lightboxPhoto && (
        <div
          onClick={() => setLightboxPhoto(null)}
          style={{
            position: "fixed", top: 0, left: 0, right: 0, bottom: 0,
            background: "rgba(0,0,0,0.88)", zIndex: 1000,
            display: "flex", alignItems: "center", justifyContent: "center",
            padding: 20, cursor: "zoom-out",
          }}
        >
          <img src={lightboxPhoto} alt="Büyütülmüş fotoğraf" style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain", borderRadius: 6 }} />
          <button
            onClick={() => setLightboxPhoto(null)}
            style={{ position: "fixed", top: 16, right: 20, background: "none", border: "none", color: "#fff", fontSize: 30, cursor: "pointer", lineHeight: 1 }}
          >
            ×
          </button>
        </div>
      )}
    </Card>
  );
}

function DriverShipments({ shipments, updateShipment, onPhotoPick, onCreate }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ title: "Nakliye", musteri: "", from: "", to: "", equipment: "", plate: "", date: "", status: "Tamamlandı" });
  const [error, setError] = useState("");
  const statusOrder = ["Bekliyor", "Yolda", "Tamamlandı"];

  function submit() {
    if (!form.from.trim() || !form.to.trim()) {
      setError("Nereden ve nereye alanları zorunlu.");
      return;
    }
    if (!form.plate.trim()) {
      setError("Kullanılan aracın plakasını girin.");
      return;
    }
    onCreate({
      title: form.title || "Nakliye",
      musteri: form.musteri.trim(),
      from: form.from.trim(),
      to: form.to.trim(),
      equipment: form.equipment.trim(),
      plate: form.plate.trim().toUpperCase(),
      date: form.date || new Date().toISOString().slice(0, 10),
      status: form.status,
      notes: "",
    });
    setForm({ title: "Nakliye", musteri: "", from: "", to: "", equipment: "", plate: "", date: "", status: "Tamamlandı" });
    setError("");
    setShowForm(false);
  }

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div style={{ fontSize: 18, fontWeight: 700 }}>Görevlerim</div>
        <PrimaryButton onClick={() => setShowForm((v) => !v)}>{showForm ? "Vazgeç" : "+ Nakliye Bildir"}</PrimaryButton>
      </div>

      {showForm && (
        <Card style={{ marginBottom: 18 }}>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="İşlem Türü">
                <select style={inputStyle} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })}>
                  {SHIPMENT_TYPES.map((t) => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Müşteri / Firma">
                <input style={inputStyle} value={form.musteri} onChange={(e) => setForm({ ...form, musteri: e.target.value })} placeholder="Örn. ABC İnşaat Ltd. Şti." />
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Nereden">
                <input style={inputStyle} value={form.from} onChange={(e) => setForm({ ...form, from: e.target.value })} placeholder="Örn. İzmir Şantiye" />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Nereye">
                <input style={inputStyle} value={form.to} onChange={(e) => setForm({ ...form, to: e.target.value })} placeholder="Örn. Manisa OSB" />
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Ekipman / iş makinesi">
                <input style={inputStyle} value={form.equipment} onChange={(e) => setForm({ ...form, equipment: e.target.value })} placeholder="Örn. Lastikli yükleyici" />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Aracın plakası">
                <input style={inputStyle} value={form.plate} onChange={(e) => setForm({ ...form, plate: e.target.value })} placeholder="Örn. 35 ABC 123" />
              </Field>
            </div>
          </div>
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Tarih">
                <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Durum">
                <select style={inputStyle} value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>
                  {statusOrder.map((st) => (
                    <option key={st} value={st}>{st}</option>
                  ))}
                </select>
              </Field>
            </div>
          </div>
          {error && <div style={{ color: RED, fontSize: 13, marginBottom: 10 }}>{error}</div>}
          <PrimaryButton onClick={submit}>Bildir</PrimaryButton>
        </Card>
      )}

      {shipments.length === 0 ? (
        <EmptyState text="Size atanmış bir nakliye görevi yok. Kendiniz de bir nakliye bildirebilirsiniz." />
      ) : (
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      {shipments.map((s) => (
        <ShipmentCard
          key={s.id}
          s={s}
          footer={
            <div style={{ marginTop: 12, display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              <select
                style={{ ...inputStyle, width: "auto", padding: "8px 10px" }}
                value={s.status}
                onChange={(e) => updateShipment(s.id, { status: e.target.value })}
              >
                {statusOrder.map((st) => (
                  <option key={st} value={st}>{st}</option>
                ))}
              </select>
              <GhostButton onClick={() => onPhotoPick(s.id)}>Fotoğraf ekle</GhostButton>
            </div>
          }
        />
      ))}
      </div>
      )}
    </div>
  );
}

function EmployeeList({ data, persist }) {
  const [name, setName] = useState("");
  const [machineName, setMachineName] = useState("");

  function add() {
    const n = name.trim();
    if (!n || data.employees.includes(n)) return;
    persist({ ...data, employees: [...data.employees, n] });
    setName("");
  }

  function addMachine() {
    const m = machineName.trim();
    if (!m || (data.machines || []).includes(m)) return;
    persist({ ...data, machines: [...(data.machines || []), m] });
    setMachineName("");
  }

  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Çalışanlar</div>
      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 8 }}>
          <input style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="Yeni çalışan adı" onKeyDown={(e) => e.key === "Enter" && add()} />
          <PrimaryButton onClick={add}>Ekle</PrimaryButton>
        </div>
      </Card>
      {data.employees.length === 0 ? (
        <EmptyState text="Henüz çalışan eklenmedi." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {data.employees.map((e) => (
            <Card key={e} style={{ padding: "10px 16px", fontWeight: 600 }}>{e}</Card>
          ))}
        </div>
      )}

      <div style={{ fontSize: 18, fontWeight: 700, margin: "28px 0 14px" }}>İş Makineleri</div>
      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 8 }}>
          <input style={inputStyle} value={machineName} onChange={(e) => setMachineName(e.target.value)} placeholder="Örn. Telehandler (Makina 1)" onKeyDown={(e) => e.key === "Enter" && addMachine()} />
          <PrimaryButton onClick={addMachine}>Ekle</PrimaryButton>
        </div>
      </Card>
      {(data.machines || []).length === 0 ? (
        <EmptyState text="Henüz iş makinesi eklenmedi." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {data.machines.map((m) => (
            <Card key={m} style={{ padding: "10px 16px", fontWeight: 600 }}>{m}</Card>
          ))}
        </div>
      )}
    </div>
  );
}

function ManagerLeaves({ data, persist }) {
  function setStatus(id, status) {
    persist({ ...data, leaveRequests: data.leaveRequests.map((l) => (l.id === id ? { ...l, status } : l)) });
  }
  if (data.leaveRequests.length === 0) {
    return <EmptyState text="Henüz izin veya rapor talebi yok." />;
  }
  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>İzin / Rapor Talepleri</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {data.leaveRequests.map((l) => (
          <Card key={l.id}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div>
                <div style={{ fontWeight: 700 }}>{l.employeeName} · {l.type}</div>
                <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>{l.startDate} - {l.endDate}</div>
                {l.reason && <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>{l.reason}</div>}
              </div>
              <Badge text={l.status} color={LEAVE_STATUS_COLORS[l.status]} />
            </div>
            {l.status === "Beklemede" && (
              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <PrimaryButton style={{ background: GREEN }} onClick={() => setStatus(l.id, "Onaylandı")}>Onayla</PrimaryButton>
                <GhostButton onClick={() => setStatus(l.id, "Reddedildi")}>Reddet</GhostButton>
              </div>
            )}
          </Card>
        ))}
      </div>
    </div>
  );
}

function DriverLeaves({ leaves, onSubmit }) {
  const [form, setForm] = useState({ type: "İzin", startDate: "", endDate: "", reason: "" });
  const [error, setError] = useState("");

  function submit() {
    if (!form.startDate || !form.endDate) {
      setError("Başlangıç ve bitiş tarihi zorunlu.");
      return;
    }
    onSubmit(form);
    setForm({ type: "İzin", startDate: "", endDate: "", reason: "" });
    setError("");
  }

  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>İzin / Rapor Bildir</div>
      <Card style={{ marginBottom: 18 }}>
        <Field label="Tür">
          <select style={inputStyle} value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
            <option value="İzin">İzin</option>
            <option value="Rapor">Rapor</option>
          </select>
        </Field>
        <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <Field label="Başlangıç">
              <input type="date" style={inputStyle} value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} />
            </Field>
          </div>
          <div style={{ flex: 1 }}>
            <Field label="Bitiş">
              <input type="date" style={inputStyle} value={form.endDate} onChange={(e) => setForm({ ...form, endDate: e.target.value })} />
            </Field>
          </div>
        </div>
        <Field label="Açıklama (opsiyonel)">
          <input style={inputStyle} value={form.reason} onChange={(e) => setForm({ ...form, reason: e.target.value })} placeholder="Kısa not" />
        </Field>
        {error && <div style={{ color: RED, fontSize: 13, marginBottom: 10 }}>{error}</div>}
        <PrimaryButton onClick={submit}>Talebi gönder</PrimaryButton>
      </Card>

      <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 10 }}>Geçmiş taleplerim</div>
      {leaves.length === 0 ? (
        <EmptyState text="Henüz bir talebiniz yok." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {leaves.map((l) => (
            <Card key={l.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontWeight: 600 }}>{l.type}</div>
                <div style={{ fontSize: 13, color: MUTED }}>{l.startDate} - {l.endDate}</div>
              </div>
              <Badge text={l.status} color={LEAVE_STATUS_COLORS[l.status]} />
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function DriverOvertime({ reports, onSubmit }) {
  const [form, setForm] = useState({ date: new Date().toISOString().slice(0, 10), hours: "", note: "" });
  const [error, setError] = useState("");

  function submit() {
    if (!form.date || !form.hours || isNaN(Number(form.hours)) || Number(form.hours) <= 0) {
      setError("Tarih ve geçerli bir fazla mesai saati girin.");
      return;
    }
    onSubmit({ date: form.date, hours: Number(form.hours), note: form.note.trim() });
    setForm({ date: new Date().toISOString().slice(0, 10), hours: "", note: "" });
    setError("");
  }

  const totalHours = reports.reduce((sum, r) => sum + Number(r.hours || 0), 0);

  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Günlük Fazla Mesai Bildir</div>
      <Card style={{ marginBottom: 18 }}>
        <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <Field label="Tarih">
              <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
            </Field>
          </div>
          <div style={{ flex: 1 }}>
            <Field label="Fazla mesai (saat)">
              <input type="number" min="0" step="0.5" style={inputStyle} value={form.hours} onChange={(e) => setForm({ ...form, hours: e.target.value })} placeholder="Örn. 2.5" />
            </Field>
          </div>
        </div>
        <Field label="Açıklama (opsiyonel)">
          <input style={inputStyle} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="Örn. Geç teslimat nedeniyle" />
        </Field>
        {error && <div style={{ color: RED, fontSize: 13, marginBottom: 10 }}>{error}</div>}
        <PrimaryButton onClick={submit}>Bildir</PrimaryButton>
      </Card>

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 15, fontWeight: 700 }}>Geçmiş bildirimlerim</div>
        <div style={{ fontSize: 13, color: MUTED }}>Toplam: {totalHours} saat</div>
      </div>
      {reports.length === 0 ? (
        <EmptyState text="Henüz fazla mesai bildiriminiz yok." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {reports.map((r) => (
            <Card key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div>
                <div style={{ fontWeight: 600 }}>{r.date}</div>
                {r.note && <div style={{ fontSize: 13, color: MUTED }}>{r.note}</div>}
              </div>
              <Badge text={`${r.hours} saat`} color={{ bg: "#FCE3E4", text: BRAND_DARK }} />
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function ManagerOvertime({ data }) {
  const reports = [...data.overtimeReports].sort((a, b) => (a.date < b.date ? 1 : -1));
  const totalsByEmployee = reports.reduce((acc, r) => {
    acc[r.employeeName] = (acc[r.employeeName] || 0) + Number(r.hours || 0);
    return acc;
  }, {});

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <div style={{ fontSize: 18, fontWeight: 700 }}>Fazla Mesai Raporları</div>
        <PrimaryButton onClick={() => window.print()} disabled={reports.length === 0}>PDF olarak kaydet</PrimaryButton>
      </div>

      {reports.length === 0 ? (
        <EmptyState text="Henüz fazla mesai bildirimi yok." />
      ) : (
        <>
          <Card style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 8, color: MUTED }}>ÇALIŞAN BAŞINA TOPLAM</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {Object.entries(totalsByEmployee).map(([name, total]) => (
                <div key={name} style={{ display: "flex", justifyContent: "space-between", fontSize: 14 }}>
                  <span>{name}</span>
                  <span style={{ fontWeight: 700 }}>{total} saat</span>
                </div>
              ))}
            </div>
          </Card>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {reports.map((r) => (
              <Card key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontWeight: 700 }}>{r.employeeName}</div>
                  <div style={{ fontSize: 13, color: MUTED }}>{r.date}{r.note ? ` · ${r.note}` : ""}</div>
                </div>
                <Badge text={`${r.hours} saat`} color={{ bg: "#FCE3E4", text: BRAND_DARK }} />
              </Card>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div style={{ textAlign: "center", padding: "40px 20px", color: MUTED, fontSize: 14, border: `1px dashed ${BORDER}`, borderRadius: 12 }}>
      {text}
    </div>
  );
}

function exportWorkReportsToExcel(reports) {
  const rows = reports.map((r) => ({
    "Tür": r.type === "aylik" ? "Aylık Puantaj" : "Günlük",
    "Operatör": r.employeeName || "",
    "Makina": r.machine || "",
    "Firma": r.firma || "",
    "Tarih": r.date || "",
    "Garaj Çıkış": r.garajCikis || "",
    "Garaj Giriş": r.garajGiris || "",
    "İşe Başlama": r.iseBaslama || "",
    "Bitiş": r.bitis || "",
    "Yol": r.yol || "",
    "Çalışma Saat": r.calismaSaat || "",
    "Fazla Çalışma": r.fazlaCalisma || "",
    "Saat Ücreti": r.saatUcreti || "",
    "Yetkili": r.yetkiliAdi || "",
    "Yetkili Tel": r.yetkiliTel || "",
    "Açıklama": r.aciklama || "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = new Array(16).fill({ wch: 15 });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Çalışma Formu");
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `calisma-formu-${dateStr}.xlsx`);
}

function DriverWorkForm({ machines, reports, onSubmit }) {
  const [formType, setFormType] = useState("gunluk");
  const blank = {
    machine: machines[0] || "",
    firma: "",
    date: new Date().toISOString().slice(0, 10),
    garajCikis: "",
    garajGiris: "",
    iseBaslama: "",
    bitis: "",
    yol: "",
    calismaSaat: "",
    fazlaCalisma: "",
    saatUcreti: "",
    yetkiliAdi: "",
    yetkiliTel: "",
    aciklama: "",
  };
  const [form, setForm] = useState(blank);
  const [error, setError] = useState("");

  function submit() {
    if (!form.machine || !form.firma.trim() || !form.date || !form.iseBaslama || !form.bitis) {
      setError("Makina, firma, tarih, işe başlama ve bitiş saati zorunlu.");
      return;
    }
    onSubmit({ ...form, type: formType });
    setForm({ ...blank, machine: form.machine, firma: formType === "aylik" ? form.firma : "" });
    setError("");
  }

  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>İş Makinesi Çalışma Formu</div>

      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <GhostButton
          onClick={() => setFormType("gunluk")}
          style={{ flex: 1, background: formType === "gunluk" ? "#FCE3E4" : "transparent", borderColor: formType === "gunluk" ? BRAND : BORDER, color: formType === "gunluk" ? BRAND_DARK : INK }}
        >
          Günlük Bildirim
        </GhostButton>
        <GhostButton
          onClick={() => setFormType("aylik")}
          style={{ flex: 1, background: formType === "aylik" ? "#E4ECF1" : "transparent", borderColor: formType === "aylik" ? STEEL : BORDER, color: formType === "aylik" ? STEEL : INK }}
        >
          Aylık Puantaj (günlük satır)
        </GhostButton>
      </div>

      <Card style={{ marginBottom: 18 }}>
        {machines.length === 0 && (
          <div style={{ fontSize: 13, color: RED, marginBottom: 10 }}>
            Henüz tanımlı iş makinesi yok — yöneticiden Çalışanlar sekmesinde eklemesini isteyin.
          </div>
        )}
        <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <Field label="Makina">
              <select style={inputStyle} value={form.machine} onChange={(e) => setForm({ ...form, machine: e.target.value })}>
                <option value="">Seçin</option>
                {machines.map((m) => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </select>
            </Field>
          </div>
          <div style={{ flex: 1 }}>
            <Field label="Tarih">
              <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
            </Field>
          </div>
        </div>
        <Field label="İşin Yapıldığı Firma">
          <input style={inputStyle} value={form.firma} onChange={(e) => setForm({ ...form, firma: e.target.value })} placeholder="Örn. ABC İnşaat Ltd. Şti." />
        </Field>

        {formType === "gunluk" && (
          <div style={{ display: "flex", gap: 12 }}>
            <div style={{ flex: 1 }}>
              <Field label="Garaj Çıkış Saati">
                <input type="time" style={inputStyle} value={form.garajCikis} onChange={(e) => setForm({ ...form, garajCikis: e.target.value })} />
              </Field>
            </div>
            <div style={{ flex: 1 }}>
              <Field label="Garaj Giriş Saati">
                <input type="time" style={inputStyle} value={form.garajGiris} onChange={(e) => setForm({ ...form, garajGiris: e.target.value })} />
              </Field>
            </div>
          </div>
        )}

        <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <Field label="İşe Başlama Saati">
              <input type="time" style={inputStyle} value={form.iseBaslama} onChange={(e) => setForm({ ...form, iseBaslama: e.target.value })} />
            </Field>
          </div>
          <div style={{ flex: 1 }}>
            <Field label="Bitiş Saati">
              <input type="time" style={inputStyle} value={form.bitis} onChange={(e) => setForm({ ...form, bitis: e.target.value })} />
            </Field>
          </div>
        </div>

        {formType === "gunluk" && (
          <Field label="Yol (güzergah)">
            <input style={inputStyle} value={form.yol} onChange={(e) => setForm({ ...form, yol: e.target.value })} placeholder="Örn. İzmir - Manisa" />
          </Field>
        )}

        <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <Field label="Çalışma Saat (toplam)">
              <input type="number" min="0" step="0.5" style={inputStyle} value={form.calismaSaat} onChange={(e) => setForm({ ...form, calismaSaat: e.target.value })} placeholder="Örn. 8" />
            </Field>
          </div>
          {formType === "aylik" ? (
            <div style={{ flex: 1 }}>
              <Field label="Fazla Çalışma (saat)">
                <input type="number" min="0" step="0.5" style={inputStyle} value={form.fazlaCalisma} onChange={(e) => setForm({ ...form, fazlaCalisma: e.target.value })} placeholder="Örn. 1.5" />
              </Field>
            </div>
          ) : (
            <div style={{ flex: 1 }}>
              <Field label="Saat Ücreti">
                <input style={inputStyle} value={form.saatUcreti} onChange={(e) => setForm({ ...form, saatUcreti: e.target.value })} placeholder="Örn. 500 TL" />
              </Field>
            </div>
          )}
        </div>

        <div style={{ display: "flex", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <Field label="Yetkili Adı">
              <input style={inputStyle} value={form.yetkiliAdi} onChange={(e) => setForm({ ...form, yetkiliAdi: e.target.value })} placeholder="Firma yetkilisi" />
            </Field>
          </div>
          {formType === "gunluk" && (
            <div style={{ flex: 1 }}>
              <Field label="Yetkili Tel">
                <input style={inputStyle} value={form.yetkiliTel} onChange={(e) => setForm({ ...form, yetkiliTel: e.target.value })} placeholder="Telefon" />
              </Field>
            </div>
          )}
        </div>

        <Field label="Açıklama (opsiyonel)">
          <input style={inputStyle} value={form.aciklama} onChange={(e) => setForm({ ...form, aciklama: e.target.value })} placeholder="Kısa not" />
        </Field>

        {error && <div style={{ color: RED, fontSize: 13, marginBottom: 10 }}>{error}</div>}
        <PrimaryButton onClick={submit}>Bildir</PrimaryButton>
      </Card>

      <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 10 }}>Geçmiş bildirimlerim</div>
      {reports.length === 0 ? (
        <EmptyState text="Henüz bir çalışma formu bildirmediniz." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {reports.map((r) => (
            <Card key={r.id}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <div style={{ fontWeight: 700 }}>{r.machine} · {r.firma}</div>
                  <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>{r.date} · {r.iseBaslama}-{r.bitis}</div>
                </div>
                <Badge text={r.type === "aylik" ? "Aylık" : "Günlük"} color={r.type === "aylik" ? { bg: "#E4ECF1", text: STEEL } : { bg: "#FCE3E4", text: BRAND_DARK }} />
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function ManagerWorkReports({ data }) {
  const [typeFilter, setTypeFilter] = useState("");
  const [machineFilter, setMachineFilter] = useState("");
  const [firmaFilter, setFirmaFilter] = useState("");
  const [driverFilter, setDriverFilter] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");

  const filtered = [...data.workReports]
    .filter((r) => (typeFilter ? r.type === typeFilter : true))
    .filter((r) => (machineFilter ? r.machine === machineFilter : true))
    .filter((r) => (driverFilter ? r.employeeName === driverFilter : true))
    .filter((r) => (firmaFilter ? (r.firma || "").toLowerCase().includes(firmaFilter.trim().toLowerCase()) : true))
    .filter((r) => (dateFrom ? r.date >= dateFrom : true))
    .filter((r) => (dateTo ? r.date <= dateTo : true))
    .sort((a, b) => (a.date < b.date ? 1 : -1));

  const hasFilter = typeFilter || machineFilter || firmaFilter || driverFilter || dateFrom || dateTo;

  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Çalışma Formu Raporları</div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: MUTED, marginBottom: 10 }}>FİLTRELE</div>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
          <div style={{ flex: 1, minWidth: 130 }}>
            <Field label="Tür">
              <select style={inputStyle} value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)}>
                <option value="">Tümü</option>
                <option value="gunluk">Günlük</option>
                <option value="aylik">Aylık</option>
              </select>
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <Field label="Makina">
              <select style={inputStyle} value={machineFilter} onChange={(e) => setMachineFilter(e.target.value)}>
                <option value="">Tümü</option>
                {(data.machines || []).map((m) => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </select>
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <Field label="Operatör">
              <select style={inputStyle} value={driverFilter} onChange={(e) => setDriverFilter(e.target.value)}>
                <option value="">Tümü</option>
                {data.employees.map((e) => (
                  <option key={e} value={e}>{e}</option>
                ))}
              </select>
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <Field label="Firma">
              <input style={inputStyle} value={firmaFilter} onChange={(e) => setFirmaFilter(e.target.value)} placeholder="Firma adı" />
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <Field label="Başlangıç">
              <input type="date" style={inputStyle} value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} />
            </Field>
          </div>
          <div style={{ flex: 1, minWidth: 130 }}>
            <Field label="Bitiş">
              <input type="date" style={inputStyle} value={dateTo} onChange={(e) => setDateTo(e.target.value)} />
            </Field>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
          {hasFilter && (
            <GhostButton onClick={() => { setTypeFilter(""); setMachineFilter(""); setFirmaFilter(""); setDriverFilter(""); setDateFrom(""); setDateTo(""); }}>
              Filtreleri Temizle
            </GhostButton>
          )}
          <PrimaryButton disabled={filtered.length === 0} onClick={() => exportWorkReportsToExcel(filtered)}>
            Excel'e Aktar ({filtered.length})
          </PrimaryButton>
        </div>
      </Card>

      {filtered.length === 0 ? (
        <EmptyState text={data.workReports.length === 0 ? "Henüz çalışma formu bildirimi yok." : "Seçilen filtrelere uyan kayıt yok."} />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {filtered.map((r) => (
            <Card key={r.id}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <div style={{ fontWeight: 700 }}>{r.machine} · {r.firma}</div>
                  <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>{r.employeeName} · {r.date} · {r.iseBaslama}-{r.bitis}</div>
                  {r.calismaSaat && <div style={{ fontSize: 13, color: MUTED, marginTop: 2 }}>Çalışma: {r.calismaSaat} saat{r.fazlaCalisma ? ` · Fazla: ${r.fazlaCalisma} saat` : ""}{r.saatUcreti ? ` · Ücret: ${r.saatUcreti}` : ""}</div>}
                </div>
                <Badge text={r.type === "aylik" ? "Aylık" : "Günlük"} color={r.type === "aylik" ? { bg: "#E4ECF1", text: STEEL } : { bg: "#FCE3E4", text: BRAND_DARK }} />
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}

function getExpiringItems(vehicles) {
  const items = [];
  vehicles.forEach((v) => {
    DOC_TYPES.forEach((dt) => {
      const doc = (v.documents || {})[dt.key] || {};
      const d = daysUntil(doc.expiryDate);
      if (d !== null && d <= 10) {
        items.push({ ownerName: v.name, docLabel: dt.label, days: d, type: "arac" });
      }
    });
  });
  return items;
}

function getExpiringEmployeeItems(employeeDocs) {
  const items = [];
  Object.entries(employeeDocs || {}).forEach(([name, docs]) => {
    EMPLOYEE_DOC_TYPES.forEach((dt) => {
      const doc = (docs || {})[dt.key] || {};
      const d = daysUntil(doc.expiryDate);
      if (d !== null && d <= 10) {
        items.push({ ownerName: name, docLabel: dt.label, days: d, type: "personel" });
      }
    });
  });
  return items;
}

function ExpiryAlertBanner({ data, onGoTo }) {
  const items = [
    ...getExpiringItems(data.vehicles || []),
    ...getExpiringEmployeeItems(data.employeeDocs || {}),
  ].sort((a, b) => a.days - b.days);

  if (items.length === 0) return null;
  const hasVehicle = items.some((i) => i.type === "arac");
  const hasPersonel = items.some((i) => i.type === "personel");

  return (
    <div style={{ background: "#F6E2E0", borderBottom: `1px solid #E3B8B2`, padding: "10px 20px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <div style={{ fontSize: 13, color: RED, fontWeight: 700 }}>
          ⚠ {items.length} belge süresi doluyor / doldu
        </div>
        <div style={{ display: "flex", gap: 12 }}>
          {hasVehicle && (
            <button onClick={() => onGoTo("araclar")} style={{ border: "none", background: "none", color: RED, fontWeight: 700, fontSize: 13, cursor: "pointer", textDecoration: "underline" }}>
              Araçları Gör
            </button>
          )}
          {hasPersonel && (
            <button onClick={() => onGoTo("personelbelgeleri")} style={{ border: "none", background: "none", color: RED, fontWeight: 700, fontSize: 13, cursor: "pointer", textDecoration: "underline" }}>
              Personel Belgelerini Gör
            </button>
          )}
        </div>
      </div>
      <div style={{ marginTop: 6, display: "flex", flexDirection: "column", gap: 2 }}>
        {items.slice(0, 4).map((it, i) => (
          <div key={i} style={{ fontSize: 12, color: RED }}>
            {it.ownerName} · {it.docLabel}: {it.days < 0 ? `${Math.abs(it.days)} gün önce doldu` : `${it.days} gün kaldı`}
          </div>
        ))}
        {items.length > 4 && <div style={{ fontSize: 12, color: RED }}>+ {items.length - 4} tane daha</div>}
      </div>
    </div>
  );
}

function VehiclesSection({ data, persist }) {
  const [showAddForm, setShowAddForm] = useState(false);
  const [newVehicleName, setNewVehicleName] = useState("");
  const [selectedId, setSelectedId] = useState(null);
  const vehicles = data.vehicles || [];
  const selected = vehicles.find((v) => v.id === selectedId);

  function addVehicle() {
    const n = newVehicleName.trim();
    if (!n) return;
    const vehicle = { id: uid(), name: n, documents: {} };
    persist({ ...data, vehicles: [...vehicles, vehicle] });
    setNewVehicleName("");
    setShowAddForm(false);
  }

  function updateVehicleDoc(vehicleId, docKey, patch) {
    const next = {
      ...data,
      vehicles: vehicles.map((v) =>
        v.id === vehicleId
          ? { ...v, documents: { ...v.documents, [docKey]: { ...(v.documents[docKey] || {}), ...patch } } }
          : v
      ),
    };
    persist(next);
  }

  function addMaintenanceEntry(vehicleId, entry) {
    const next = {
      ...data,
      vehicles: vehicles.map((v) =>
        v.id === vehicleId
          ? { ...v, maintenanceHistory: [{ ...entry, id: uid(), createdAt: new Date().toISOString() }, ...(v.maintenanceHistory || [])] }
          : v
      ),
    };
    persist(next);
  }

  if (selected) {
    return (
      <VehicleDetail
        vehicle={selected}
        onBack={() => setSelectedId(null)}
        onUpdateDoc={(docKey, patch) => updateVehicleDoc(selected.id, docKey, patch)}
        onAddMaintenance={(entry) => addMaintenanceEntry(selected.id, entry)}
      />
    );
  }

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 18, fontWeight: 700 }}>Araçlar</div>
        <div style={{ display: "flex", gap: 8 }}>
          <GhostButton
            onClick={() => exportMaintenanceToExcel(vehicles.flatMap((v) => (v.maintenanceHistory || []).map((m) => ({ ...m, vehicleName: v.name }))))}
          >
            Bakım Raporu (Tüm Filo)
          </GhostButton>
          <PrimaryButton onClick={() => setShowAddForm((v) => !v)}>{showAddForm ? "Vazgeç" : "+ Araç Ekle"}</PrimaryButton>
        </div>
      </div>

      {showAddForm && (
        <Card style={{ marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 8 }}>
            <input
              style={inputStyle}
              value={newVehicleName}
              onChange={(e) => setNewVehicleName(e.target.value)}
              placeholder="Örn. 35 ABC 123 - Çekici"
              onKeyDown={(e) => e.key === "Enter" && addVehicle()}
            />
            <PrimaryButton onClick={addVehicle}>Ekle</PrimaryButton>
          </div>
        </Card>
      )}

      {vehicles.length === 0 ? (
        <EmptyState text="Henüz araç eklenmedi." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {vehicles.map((v) => {
            const items = getExpiringItems([v]);
            return (
              <Card key={v.id} style={{ cursor: "pointer" }} onClick={() => setSelectedId(v.id)}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{v.name}</div>
                  {items.length > 0 ? (
                    <Badge text={`${items.length} uyarı`} color={{ bg: "#F6E2E0", text: RED }} />
                  ) : (
                    <Badge text="Belgeler güncel" color={{ bg: "#E7F0E3", text: GREEN }} />
                  )}
                </div>
                <div style={{ fontSize: 12, color: MUTED, marginTop: 4 }}>Belgeleri görmek için tıklayın →</div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

function VehicleDetail({ vehicle, onBack, onUpdateDoc, onAddMaintenance }) {
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <GhostButton onClick={onBack}>← Geri</GhostButton>
        <div style={{ fontSize: 18, fontWeight: 700 }}>{vehicle.name}</div>
      </div>
      <MaintenanceSection vehicle={vehicle} onAdd={onAddMaintenance} />
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {DOC_TYPES.map((dt) => (
          <DocumentRow
            key={dt.key}
            storageKey={`vehicle-file:${vehicle.id}:${dt.key}`}
            label={dt.label}
            doc={(vehicle.documents || {})[dt.key] || {}}
            onUpdate={(patch) => onUpdateDoc(dt.key, patch)}
          />
        ))}
      </div>
    </div>
  );
}

function exportMaintenanceToExcel(entries) {
  const rows = [...entries]
    .sort((a, b) => (a.date < b.date ? 1 : -1))
    .map((m) => ({
      "Araç": m.vehicleName || "",
      "Tarih": m.date || "",
      "Km": m.km || "",
      "Bakım Detayı": m.detail || "",
    }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{ wch: 22 }, { wch: 14 }, { wch: 12 }, { wch: 40 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Bakım Kayıtları");
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `bakim-raporu-${dateStr}.xlsx`);
}

function MaintenanceSection({ vehicle, onAdd }) {
  const [km, setKm] = useState("");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [detail, setDetail] = useState("");
  const [error, setError] = useState("");
  const history = vehicle.maintenanceHistory || [];

  function submit() {
    if (!date || !km) {
      setError("Tarih ve km bilgisi zorunlu.");
      return;
    }
    onAdd({ km, date, detail: detail.trim() });
    setKm("");
    setDetail("");
    setError("");
  }

  return (
    <Card style={{ marginBottom: 10 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <div style={{ fontSize: 13, fontWeight: 700, color: MUTED }}>BAKIM KAYITLARI</div>
        {history.length > 0 && (
          <GhostButton
            onClick={() => exportMaintenanceToExcel(history.map((m) => ({ ...m, vehicleName: vehicle.name })))}
          >
            Excel'e Aktar
          </GhostButton>
        )}
      </div>

      <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 130 }}>
          <Field label="Bakım Km">
            <input type="number" min="0" style={inputStyle} value={km} onChange={(e) => setKm(e.target.value)} placeholder="Örn. 125000" />
          </Field>
        </div>
        <div style={{ flex: 1, minWidth: 130 }}>
          <Field label="Bakım Tarihi">
            <input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} />
          </Field>
        </div>
      </div>
      <Field label="Yapılan Bakım Detayı">
        <input style={inputStyle} value={detail} onChange={(e) => setDetail(e.target.value)} placeholder="Örn. Yağ, filtre değişimi, fren balatası" />
      </Field>
      {error && <div style={{ color: RED, fontSize: 13, marginBottom: 10 }}>{error}</div>}
      <PrimaryButton onClick={submit}>Kaydet</PrimaryButton>

      {history.length > 0 && (
        <div style={{ marginTop: 16, display: "flex", flexDirection: "column", gap: 8 }}>
          {history.map((m) => (
            <div key={m.id} style={{ padding: "10px 12px", background: "#FAF9F6", borderRadius: 8, border: `1px solid ${BORDER}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, fontWeight: 700 }}>
                <span>{m.date}</span>
                <span>{m.km} km</span>
              </div>
              {m.detail && <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>{m.detail}</div>}
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}

function DocumentRow({ storageKey, label, doc, onUpdate, readOnly }) {
  const fileInputRef = useRef(null);
  const [expiryDate, setExpiryDate] = useState(doc.expiryDate || "");
  const [uploading, setUploading] = useState(false);
  const [downloadError, setDownloadError] = useState("");
  const status = docStatus(doc.expiryDate);

  async function handleFile(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setDownloadError("");
    setUploading(true);
    try {
      let payload;
      if (file.type.startsWith("image/")) {
        payload = await resizeImage(file, 1600);
      } else {
        if (file.size > 4 * 1024 * 1024) {
          setDownloadError(`Dosya çok büyük (${(file.size / (1024 * 1024)).toFixed(1)} MB). 4 MB altında bir PDF/dosya yükleyin.`);
          setUploading(false);
          return;
        }
        payload = await fileToDataUrl(file);
      }
      const approxBytes = payload.length * 0.75;
      if (approxBytes > 4.5 * 1024 * 1024) {
        setDownloadError(`Dosya çok büyük (${(approxBytes / (1024 * 1024)).toFixed(1)} MB). Daha küçük bir dosya deneyin.`);
        setUploading(false);
        return;
      }
      const res = await window.storage.set(storageKey, payload, true);
      if (res) {
        onUpdate({ fileName: file.name, hasFile: true });
      } else {
        setDownloadError("Dosya kaydedilemedi (muhtemelen boyut sınırı). Daha küçük bir dosya deneyin.");
      }
    } catch (err) {
      setDownloadError("Dosya yüklenemedi: " + (err && err.message ? err.message : "bilinmeyen hata") + ".");
    }
    setUploading(false);
  }

  async function handleDownload() {
    setDownloadError("");
    try {
      const res = await window.storage.get(storageKey, true);
      if (res && res.value) {
        downloadDataUrl(res.value, doc.fileName || label);
      } else {
        setDownloadError("Dosya bulunamadı.");
      }
    } catch (err) {
      setDownloadError("Dosya indirilemedi.");
    }
  }

  function saveExpiry() {
    onUpdate({ expiryDate });
  }

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, flexWrap: "wrap" }}>
        <div style={{ fontWeight: 700 }}>{label}</div>
        <Badge text={status.label} color={status.color} />
      </div>
      <div style={{ display: "flex", gap: 12, marginTop: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
        {readOnly ? (
          <div style={{ fontSize: 13, color: MUTED }}>
            Bitiş Tarihi: {doc.expiryDate || "Girilmedi"}
          </div>
        ) : (
          <div style={{ flex: 1, minWidth: 160 }}>
            <Field label="Bitiş Tarihi">
              <input type="date" style={inputStyle} value={expiryDate} onChange={(e) => setExpiryDate(e.target.value)} onBlur={saveExpiry} />
            </Field>
          </div>
        )}
        {!readOnly && (
          <>
            <input type="file" accept="application/pdf,image/*" ref={fileInputRef} onChange={handleFile} style={{ display: "none" }} />
            <GhostButton style={{ marginBottom: 14 }} onClick={() => fileInputRef.current?.click()}>
              {uploading ? "Yükleniyor..." : doc.hasFile ? "Belgeyi Değiştir" : "Belge Yükle"}
            </GhostButton>
          </>
        )}
        {doc.hasFile && (
          <PrimaryButton style={readOnly ? {} : { marginBottom: 14 }} onClick={handleDownload}>
            İndir
          </PrimaryButton>
        )}
      </div>
      {doc.fileName && <div style={{ fontSize: 12, color: MUTED, marginTop: 2 }}>Dosya: {doc.fileName}</div>}
      {downloadError && <div style={{ fontSize: 12, color: RED, marginTop: 4 }}>{downloadError}</div>}
    </Card>
  );
}

function EmployeeDocsSection({ data, persist }) {
  const [selectedName, setSelectedName] = useState(null);
  const employees = data.employees || [];
  const employeeDocs = data.employeeDocs || {};

  function updateDoc(name, docKey, patch) {
    const current = employeeDocs[name] || {};
    const next = {
      ...data,
      employeeDocs: {
        ...employeeDocs,
        [name]: { ...current, [docKey]: { ...(current[docKey] || {}), ...patch } },
      },
    };
    persist(next);
  }

  if (selectedName) {
    return (
      <div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
          <GhostButton onClick={() => setSelectedName(null)}>← Geri</GhostButton>
          <div style={{ fontSize: 18, fontWeight: 700 }}>{selectedName}</div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {EMPLOYEE_DOC_TYPES.map((dt) => (
            <DocumentRow
              key={dt.key}
              storageKey={`employee-file:${safeKeyPart(selectedName)}:${dt.key}`}
              label={dt.label}
              doc={(employeeDocs[selectedName] || {})[dt.key] || {}}
              onUpdate={(patch) => updateDoc(selectedName, dt.key, patch)}
            />
          ))}
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 14 }}>Personel Belgeleri</div>
      {employees.length === 0 ? (
        <EmptyState text="Henüz çalışan eklenmedi." />
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {employees.map((name) => {
            const items = getExpiringEmployeeItems({ [name]: employeeDocs[name] });
            return (
              <Card key={name} style={{ cursor: "pointer" }} onClick={() => setSelectedName(name)}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontWeight: 700, fontSize: 15 }}>{name}</div>
                  {items.length > 0 ? (
                    <Badge text={`${items.length} uyarı`} color={{ bg: "#F6E2E0", text: RED }} />
                  ) : (
                    <Badge text="Belgeler güncel" color={{ bg: "#E7F0E3", text: GREEN }} />
                  )}
                </div>
                <div style={{ fontSize: 12, color: MUTED, marginTop: 4 }}>Belgeleri görmek için tıklayın →</div>
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}

function MyDocuments({ employeeName, docs }) {
  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 4 }}>Belgelerim</div>
      <div style={{ fontSize: 13, color: MUTED, marginBottom: 14 }}>Bu belgeleri sadece siz görebilir ve indirebilirsiniz.</div>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {EMPLOYEE_DOC_TYPES.map((dt) => (
          <DocumentRow
            key={dt.key}
            storageKey={`employee-file:${safeKeyPart(employeeName)}:${dt.key}`}
            label={dt.label}
            doc={(docs || {})[dt.key] || {}}
            readOnly
          />
        ))}
      </div>
    </div>
  );
}
