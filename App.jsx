import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient";
import AuthLogin from "./AuthLogin";

const KIRLI_WORDMARK_SRC = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAPUAAAArCAYAAACghc46AAAEIUlEQVR4nO2dzXHbMBBGn2I3oJQgn3NSSlBKcEqgSnBKsEqQSohKiEuILr5HLbCBTJIDiKHiOGMKBD6AxL4ZjT0ekwQX+rD4WSwWz4ubDWlpgVOC+4aW+2V51O+f+nkxSVF3U/2+xSL5+y+eFze/Ez/kCfgU8X5L4BFoAq5tu7JcVrr6/VM/LxUn4Nz9POHeK4Spfd9ik/z9bxM/IDYr4CuwDrj2hKvsNmqJ6mHdfe4v/nYAjoQL3EjAu9wFuIIN8B0TdEk0wLfuM6VhxayZiqj9l2cZcO0RE3RqNrj6echdEGMaot53nxAOwGdM0CoeCa8rIxIli3qJ626HTIiBE/Q2XnGMgTSE15kRgVJFvQZ+EDZ+BidmE3Q+9oTXnTGSEme/G8Z14bY4L10rJ+BLwHVehJvu95D5i0v2wMeR9zACKE3Ue8K7bi1O0Md4xZkkLWFLTP6aXffzHlcXobPa6+76mhvYLJTS/R47fvZBJbULOiZ+1SDE63tsbJ2BEkQ9dvz8WpSYEY8d4cL2ASuGkNyibnAeOnT8dgLuMEGnZocLEQ3BglLE5BT1mPVn6D20rUFrMFFPhByiXuKij8aOt5ZY124KmKjFqEW9xnW3Y1W0hSUaxguUovbx26uI99xgnkDBmF6RzXeIUYnaxwSPDWh4DfPWaVnjtruG1p3NeYhRBJ+k9qb+/ran17GhrEQMVi9ici9pxeL+7X8xMmGiFjMXUTfEHasbcfDpjwwhcxE12Ni6RMaEmBqBzEnU5q3LwucvM8TMSdRg3roEWpyHtv3smVDMfvs8zCvSe9KGcXHKc0Bp78sc2z6F8BFbxsqKQtQ+k6ePJkvNA3V7CbW9LQdcYSi73yc0G+ZtbO1Q2HuJ7ZkuDvWYeoemVbextUNh7wfSRAoagahFfaZPl5MS89YOhb39MUhDSd3IVF/vOWa/D5i3VqKw9zWNaOpgFMUEYdHkELVf8kiNeWuHyt5DG1FFhFnVDXqudeoDmmWnqiv3AoW9G4Ztz1SIuuoDBXIGnyiWncxb9yjsPWRsrVrH3lNpo55T1E9odvBUWbGvoLD30G22islScI2MTz1dzQx97kPnVQESd/y/+zm3Q+dz23vIoe9LXFpotdDO9DvHVAEzLX/HC8z+0HkfIJF6/FN7lJlHYe8hSStanLe+ZiksBn5mXJkC6wnxKSUlbOhQBEhU1f16A4W9h6R+3mFH8iShBFErA1IMjb1XDLN37YcZJqEEUYMmQMLCGXtU9h7CFreObptCIlGKqBUBErb5oEdh7xXDhb3DHXtrXjsCpYgaNAES5q17SrP3Gee13+MaHEtYGMgt6Y13TQTRFpfwPxXeW1+OKdXvX7u938LPjPtrNvSRalM8uEFe/4vnxU3iZ1xN6oq7zNZhmL0nzYdfP//52x+sVvFToUFpDQAAAABJRU5ErkJggg==";

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
  "employeeLocations",
  "employeeLeaveEntitlements",
];

function daysBetweenInclusive(start, end) {
  if (!start || !end) return 0;
  const s = new Date(start);
  const e = new Date(end);
  const diff = Math.round((e - s) / 86400000) + 1;
  return diff > 0 ? diff : 0;
}

function computeUsedLeaveDays(leaveRequests, employeeName) {
  return leaveRequests
    .filter((l) => l.employeeName === employeeName && l.type === "İzin" && l.status === "Onaylandı")
    .reduce((sum, l) => sum + daysBetweenInclusive(l.startDate, l.endDate), 0);
}

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
      const base = { employees: [], shipments: [], leaveRequests: [], overtimeReports: [], machines: [], workReports: [], vehicles: [], employeeDocs: {}, employeeLocations: {}, employeeLeaveEntitlements: {} };
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
    const prev = data;
    setData(next);
    try {
      const changedKeys = STORAGE_KEYS.filter(
        (k) => JSON.stringify(next[k]) !== JSON.stringify(prev ? prev[k] : undefined)
      );
      const results = await Promise.all(
        changedKeys.map((k) => window.storage.set(`data:${k}`, JSON.stringify(next[k]), true))
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

  function captureLocationSilently() {
    if (role !== "calisan" || !currentName || !navigator.geolocation) return;
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        const loc = {
          lat: pos.coords.latitude,
          lng: pos.coords.longitude,
          timestamp: new Date().toISOString(),
        };
        setData((prev) => {
          if (!prev) return prev;
          const nextLocations = { ...(prev.employeeLocations || {}), [currentName]: loc };
          window.storage.set("data:employeeLocations", JSON.stringify(nextLocations), true);
          return { ...prev, employeeLocations: nextLocations };
        });
      },
      () => {
        // Konum izni verilmediyse veya alınamadıysa sessizce geç —
        // nakliye/durum bildirimini bu yüzden engellemiyoruz.
      },
      { enableHighAccuracy: true, timeout: 8000 }
    );
  }

  function updateShipment(id, patch) {
    const next = {
      ...data,
      shipments: data.shipments.map((s) => (s.id === id ? { ...s, ...patch } : s)),
    };
    persist(next);
    if (patch.status) captureLocationSilently();
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
      <div
        className="tab-scroll"
        style={{
          display: "flex",
          gap: 6,
          padding: "10px 12px",
          overflowX: "auto",
          WebkitOverflowScrolling: "touch",
          scrollbarWidth: "none",
        }}
      >
        {[
          ["nakliyeler", role === "yonetici" ? "Nakliyeler" : "Görevlerim"],
          ["izinler", "İzin / Rapor"],
          ["mesai", "Fazla Mesai"],
          ["calismaformu", "Çalışma Formu"],
          ...(role === "yonetici" ? [["calisanlar", "Çalışanlar"], ["araclar", "Araçlar"], ["personelbelgeleri", "Personel Belgeleri"], ["konumlar", "Şoför Konumları"]] : [["belgelerim", "Belgelerim"]]),
        ].map(([key, label]) => (
          <button
            key={key}
            onClick={() => setTab(key)}
            style={{
              border: tab === key ? "none" : `1px solid ${BORDER}`,
              background: tab === key ? BRAND : "#fff",
              cursor: "pointer",
              padding: "8px 14px",
              fontSize: 12.5,
              fontWeight: 700,
              whiteSpace: "nowrap",
              flexShrink: 0,
              borderRadius: 999,
              color: tab === key ? "#fff" : MUTED,
            }}
          >
            {label}
          </button>
        ))}
      </div>

      <style>{`
        .tab-scroll::-webkit-scrollbar { display: none; }
      `}</style>

      <div style={{ padding: 20, maxWidth: 760, margin: "0 auto" }}>
        {tab === "nakliyeler" && role === "yonetici" && (
          <ManagerShipments data={data} persist={persist} />
        )}
        {tab === "nakliyeler" && role === "calisan" && (
          <DriverShipments
            shipments={myShipments}
            updateShipment={updateShipment}
            onPhotoPick={handlePhotoPick}
            onCreate={(s) => {
              persist({
                ...data,
                shipments: [
                  { ...s, id: uid(), assignedTo: currentName, source: "surucu", photos: [], createdAt: new Date().toISOString() },
                  ...data.shipments,
                ],
              });
              captureLocationSilently();
            }}
          />
        )}
        {tab === "izinler" && role === "yonetici" && (
          <ManagerLeaves data={data} persist={persist} />
        )}
        {tab === "izinler" && role === "calisan" && (
          <DriverLeaves
            leaves={myLeaves}
            entitlement={(data.employeeLeaveEntitlements || {})[currentName]}
            usedDays={computeUsedLeaveDays(data.leaveRequests, currentName)}
            onSubmit={(req) => persist({ ...data, leaveRequests: [{ ...req, id: uid(), employeeName: currentName, status: "Beklemede" }, ...data.leaveRequests] })}
          />
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
        {tab === "konumlar" && role === "yonetici" && (
          <DriverLocationsSection data={data} />
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
        <img src={KIRLI_WORDMARK_SRC} alt="Kırlı" style={{ height: 40, marginBottom: 10 }} />
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
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px", borderBottom: `1px solid ${BORDER}` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <img src={KIRLI_WORDMARK_SRC} alt="Kırlı" style={{ height: 22, width: "auto" }} />
        <div>
          <div style={{ fontWeight: 800, fontSize: 13 }}>Nakliye Takip</div>
          <div style={{ fontSize: 11, color: MUTED }}>{currentName} · {role === "yonetici" ? "Yönetici" : "Çalışan"}</div>
        </div>
      </div>
      <GhostButton onClick={onSwitch} style={{ fontSize: 12, padding: "6px 10px" }}>Çıkış</GhostButton>
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
            <EmployeeLeaveRow
              key={e}
              name={e}
              entitlement={(data.employeeLeaveEntitlements || {})[e]}
              usedDays={computeUsedLeaveDays(data.leaveRequests, e)}
              onSave={(days) =>
                persist({
                  ...data,
                  employeeLeaveEntitlements: { ...(data.employeeLeaveEntitlements || {}), [e]: days },
                })
              }
            />
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

function EmployeeLeaveRow({ name, entitlement, usedDays, onSave }) {
  const [value, setValue] = useState(entitlement ?? "");
  const remaining = entitlement !== undefined && entitlement !== null && entitlement !== "" ? Number(entitlement) - usedDays : null;

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontWeight: 700 }}>{name}</div>
        <div style={{ fontSize: 12, color: MUTED }}>
          {remaining !== null ? (
            <>Kullanılan: <b style={{ color: INK }}>{usedDays}</b> gün · Kalan: <b style={{ color: remaining < 0 ? RED : GREEN }}>{remaining}</b> gün</>
          ) : (
            "Yıllık izin hakkı tanımlanmadı"
          )}
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 10, alignItems: "center" }}>
        <input
          type="number"
          min="0"
          style={{ ...inputStyle, width: 90 }}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder="Örn. 14"
        />
        <span style={{ fontSize: 12, color: MUTED }}>gün / yıl</span>
        <GhostButton onClick={() => onSave(value === "" ? null : Number(value))}>Kaydet</GhostButton>
      </div>
    </Card>
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
                <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>{l.startDate} - {l.endDate} ({daysBetweenInclusive(l.startDate, l.endDate)} gün)</div>
                {l.reason && <div style={{ fontSize: 13, color: MUTED, marginTop: 3 }}>{l.reason}</div>}
                {l.type === "İzin" && (data.employeeLeaveEntitlements || {})[l.employeeName] !== undefined && (
                  <div style={{ fontSize: 12, color: MUTED, marginTop: 3 }}>
                    Yıllık hak: {data.employeeLeaveEntitlements[l.employeeName]} gün · Kullanılan (bu dahil onaylılar): {computeUsedLeaveDays(data.leaveRequests, l.employeeName)} gün
                  </div>
                )}
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

function DriverLeaves({ leaves, entitlement, usedDays, onSubmit }) {
  const [form, setForm] = useState({ type: "İzin", startDate: "", endDate: "", reason: "" });
  const [error, setError] = useState("");
  const hasEntitlement = entitlement !== undefined && entitlement !== null;
  const remaining = hasEntitlement ? Number(entitlement) - usedDays : null;

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

      {hasEntitlement && (
        <Card style={{ marginBottom: 16, background: "#FAF9F6" }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: MUTED, marginBottom: 8 }}>YILLIK İZİN HAKKINIZ</div>
          <div style={{ display: "flex", gap: 20 }}>
            <div>
              <div style={{ fontSize: 22, fontWeight: 800 }}>{entitlement}</div>
              <div style={{ fontSize: 11, color: MUTED }}>Toplam gün</div>
            </div>
            <div>
              <div style={{ fontSize: 22, fontWeight: 800 }}>{usedDays}</div>
              <div style={{ fontSize: 11, color: MUTED }}>Kullanılan</div>
            </div>
            <div>
              <div style={{ fontSize: 22, fontWeight: 800, color: remaining < 0 ? RED : GREEN }}>{remaining}</div>
              <div style={{ fontSize: 11, color: MUTED }}>Kalan</div>
            </div>
          </div>
        </Card>
      )}

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

function DriverLocationsMap({ locations }) {
  const mapContainerRef = useRef(null);
  const mapInstanceRef = useRef(null);
  const markersRef = useRef([]);

  useEffect(() => {
    let cancelled = false;

    import("leaflet").then((L) => {
      if (cancelled) return;
      const leaflet = L.default || L;

      const icon = new leaflet.Icon({
        iconUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png",
        iconRetinaUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png",
        shadowUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png",
        iconSize: [25, 41],
        iconAnchor: [12, 41],
      });

      if (!mapInstanceRef.current) {
        mapInstanceRef.current = leaflet.map(mapContainerRef.current).setView([38.42, 27.14], 8);
        leaflet
          .tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
            attribution: "© OpenStreetMap contributors",
            maxZoom: 19,
          })
          .addTo(mapInstanceRef.current);
      }

      markersRef.current.forEach((m) => m.remove());
      markersRef.current = [];

      const entries = Object.entries(locations || {}).filter(([, loc]) => loc && loc.lat && loc.lng);

      entries.forEach(([name, loc]) => {
        const minutesAgo = loc.timestamp
          ? Math.max(0, Math.round((Date.now() - new Date(loc.timestamp).getTime()) / 60000))
          : null;
        const marker = leaflet.marker([loc.lat, loc.lng], { icon }).addTo(mapInstanceRef.current);
        marker.bindPopup(
          `<b>${name}</b><br/>${minutesAgo !== null ? minutesAgo + " dakika önce" : "Zaman bilinmiyor"}`
        );
        markersRef.current.push(marker);
      });

      if (entries.length > 0) {
        const bounds = leaflet.latLngBounds(entries.map(([, loc]) => [loc.lat, loc.lng]));
        if (bounds.isValid()) {
          mapInstanceRef.current.fitBounds(bounds, { padding: [30, 30], maxZoom: 13 });
        }
      }

      setTimeout(() => mapInstanceRef.current && mapInstanceRef.current.invalidateSize(), 100);
    });

    return () => {
      cancelled = true;
    };
  }, [locations]);

  useEffect(() => {
    return () => {
      if (mapInstanceRef.current) {
        mapInstanceRef.current.remove();
        mapInstanceRef.current = null;
      }
    };
  }, []);

  return (
    <div
      ref={mapContainerRef}
      style={{ width: "100%", height: 420, borderRadius: 12, overflow: "hidden", border: `1px solid ${BORDER}` }}
    />
  );
}

function DriverLocationsSection({ data }) {
  const locations = data.employeeLocations || {};
  const entries = Object.entries(locations).sort((a, b) => {
    const ta = a[1]?.timestamp ? new Date(a[1].timestamp).getTime() : 0;
    const tb = b[1]?.timestamp ? new Date(b[1].timestamp).getTime() : 0;
    return tb - ta;
  });

  return (
    <div>
      <div style={{ fontSize: 18, fontWeight: 700, marginBottom: 4 }}>Şoför Konumları</div>
      <div style={{ fontSize: 13, color: MUTED, marginBottom: 14 }}>
        Bir şoför nakliye bildirdiğinde veya durumunu güncellediğinde konumu otomatik olarak burada görünür.
      </div>

      {entries.length === 0 ? (
        <EmptyState text="Henüz konum paylaşımı yok." />
      ) : (
        <>
          <DriverLocationsMap locations={locations} />
          <div style={{ marginTop: 16, display: "flex", flexDirection: "column", gap: 8 }}>
            {entries.map(([name, loc]) => {
              const minutesAgo = loc.timestamp
                ? Math.max(0, Math.round((Date.now() - new Date(loc.timestamp).getTime()) / 60000))
                : null;
              return (
                <Card key={name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontWeight: 600 }}>{name}</div>
                  <div style={{ fontSize: 13, color: MUTED }}>
                    {minutesAgo !== null ? `${minutesAgo} dk önce` : "Zaman bilinmiyor"}
                  </div>
                </Card>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
}
