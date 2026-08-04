import React, { useEffect, useMemo, useState } from "react";
import {
  LayoutDashboard, Package, Users, FileText, Wallet, BookOpen,
  Plus, Trash2, Pencil, Printer, X, AlertTriangle, Check, Search, ClipboardList,
  Factory, Boxes, Lock, Unlock, Settings2, Banknote, BarChart3, TrendingUp, Award, Calculator, Truck,
} from "lucide-react";
import * as XLSX from "xlsx";
import { api } from "./api";

const todayISO = () => new Date().toISOString().slice(0, 10);
const fmtNum = (n) => Math.round(Number(n) || 0).toLocaleString("uz-UZ");
const fmt = (n) => fmtNum(n) + " so'm";
const fmtQty = (n) => (Number(n) || 0).toFixed(1);
const uid = () => Math.random().toString(36).slice(2);
const fmtDate = (d) => {
  if (!d) return "";
  const dt = new Date(d);
  return dt.toLocaleDateString("uz-UZ", { day: "2-digit", month: "2-digit", year: "numeric" });
};

function exportAllToExcel(data) {
  const {
    products, customers, invoices, payments, rawMaterials, rawMaterialBatches,
    finishedProducts, productionBatches, cashTransactions, balances,
    rawMaterialsValue, stockValue, totalDebt, totalCredit, cashBalanceSom, cashBalanceUsd, usdRate,
  } = data;

  const wb = XLSX.utils.book_new();
  const sheet = (rows, name) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);

  sheet([
    ["Bo'lim", "Summa"],
    ["Xom ashyo ombori (faol)", rawMaterialsValue],
    ["Tayyor mahsulot ombori", stockValue],
    ["Debitor (mijozlar qarzi)", totalDebt],
    ["Kreditor (ortiqcha to'langan)", totalCredit],
    ["Kassa (so'm)", cashBalanceSom],
    ["Kassa (dollar)", cashBalanceUsd],
    ["Kurs (1$)", usdRate],
    ["Jami aktivlar (so'm ekvivalent)", rawMaterialsValue + stockValue + totalDebt + cashBalanceSom + cashBalanceUsd * usdRate],
  ], "Balans");

  sheet([
    ["Nomi", "Artikul", "Birlik", "Narxi", "Qoldiq", "Qiymati", "Tannarx"],
    ...products.map((p) => [p.name, p.article || "", p.unit, p.price, p.qty, p.qty * p.price, p.cost_price || 0]),
  ], "Ombor");

  sheet([
    ["Xom ashyo", "Sana", "Kelgan", "Qolgan", "Tannarx", "Holat"],
    ...rawMaterialBatches.map((b) => {
      const m = rawMaterials.find((x) => x.id === b.raw_material_id);
      return [m?.name || "", b.date, b.qty, b.remaining_qty, b.unit_cost, b.active ? "Faol" : "Harakatsiz"];
    }),
  ], "Xom ashyo");

  sheet([
    ["Nomi", "Telefon", "Jami faktura", "Jami to'lov", "Balans"],
    ...customers.map((c) => {
      const b = balances[c.id] || { invoiced: 0, paid: 0 };
      return [c.name, c.phone || "", b.invoiced, b.paid, b.invoiced - b.paid];
    }),
  ], "Mijozlar");

  sheet([
    ["Raqam", "Sana", "Mijoz", "Summa"],
    ...invoices.map((i) => [i.number, i.date, customers.find((c) => c.id === i.customer_id)?.name || "", i.total]),
  ], "Fakturalar");

  const kassaRows = (curr) => {
    const list = cashTransactions.filter((t) => (t.currency || "som") === curr).sort((a, b) => new Date(a.date) - new Date(b.date));
    let running = 0;
    const header = curr === "usd" ? ["Sana", "Mazmuni", "Bo'lim", "Kurs", "Kirim", "Chiqim", "Qoldiq"] : ["Sana", "Mazmuni", "Bo'lim", "Kirim", "Chiqim", "Qoldiq"];
    return [
      header,
      ...list.map((t) => {
        running += t.type === "kirim" ? Number(t.amount) : -Number(t.amount);
        const row = [t.date, t.note || "", t.category || ""];
        if (curr === "usd") row.push(t.rate || "");
        row.push(t.type === "kirim" ? t.amount : "", t.type === "chiqim" ? t.amount : "", running);
        return row;
      }),
    ];
  };
  sheet(kassaRows("som"), "Kassa - som");
  sheet(kassaRows("usd"), "Kassa - dollar");

  sheet([
    ["Sana", "Mahsulot", "Miqdor", "Birlik tannarx", "Jami tannarx"],
    ...productionBatches.map((b) => [b.date, finishedProducts.find((f) => f.id === b.finished_product_id)?.name || "", b.qty, b.unit_cost, b.total_cost]),
  ], "Ishlab chiqarish");

  sheet([
    ["Sana", "Mijoz", "Summa", "Izoh"],
    ...payments.map((p) => [p.date, customers.find((c) => c.id === p.customer_id)?.name || "", p.amount, p.note || ""]),
  ], "To'lovlar");

  XLSX.writeFile(wb, `Ombor-hisobot-${todayISO()}.xlsx`);
}

export default function App() {
  const [tab, setTab] = useState("dashboard");
  const [ready, setReady] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [showCompanyModal, setShowCompanyModal] = useState(false);
  const [productionPrefill, setProductionPrefill] = useState(null); // { productId, qty }
  const [invCustomerId, setInvCustomerId] = useState("");
  const [invDate, setInvDate] = useState(todayISO());
  const [invRows, setInvRows] = useState([]);
  const [currentUser, setCurrentUser] = useState(() => {
    try { return localStorage.getItem("ombor_user") || ""; } catch (e) { return ""; }
  });
  const [showUserModal, setShowUserModal] = useState(false);
  const [unlocked, setUnlocked] = useState(() => {
    try { return sessionStorage.getItem("ombor_unlocked") === "1"; } catch (e) { return false; }
  });
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [products, setProducts] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [payments, setPayments] = useState([]);
  const [movements, setMovements] = useState([]);
  const [finishedProducts, setFinishedProducts] = useState([]);
  const [rawMaterials, setRawMaterials] = useState([]);
  const [rawMaterialBatches, setRawMaterialBatches] = useState([]);
  const [productNorms, setProductNorms] = useState([]);
  const [productionBatches, setProductionBatches] = useState([]);
  const [productionConsumptions, setProductionConsumptions] = useState([]);
  const [cashTransactions, setCashTransactions] = useState([]);
  const [cashCategories, setCashCategories] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  const [supplierPayments, setSupplierPayments] = useState([]);
  const [fixedAssets, setFixedAssets] = useState([]);
  const [settings, setSettings] = useState({ company_name: "Mening korxonam", company_phone: "" });
  const [printInvoice, setPrintInvoice] = useState(null);

  const reloadAll = async () => {
    try {
      const [p, c, inv, pay, mov, set, fp, rm, rmb, pn, pb, ct, pc, cc, sup, sp, fa] = await Promise.all([
        api.products.list(), api.customers.list(), api.invoices.list(), api.payments.list(), api.stockMovements.list(), api.settings.get(),
        api.finishedProducts.list(), api.rawMaterials.list(), api.rawMaterialBatches.list(), api.productNorms.list(), api.productionBatches.list(),
        api.cashTransactions.list(), api.productionConsumptions.list(), api.cashCategories.list(), api.suppliers.list(), api.supplierPayments.list(),
        api.fixedAssets.list(),
      ]);
      setProducts(p || []); setCustomers(c || []); setInvoices(inv || []); setPayments(pay || []); setMovements(mov || []); setSettings(set);
      setFinishedProducts(fp || []); setRawMaterials(rm || []); setRawMaterialBatches(rmb || []); setProductNorms(pn || []); setProductionBatches(pb || []);
      setCashTransactions(ct || []); setProductionConsumptions(pc || []); setCashCategories(cc || []);
      setSuppliers(sup || []); setSupplierPayments(sp || []); setFixedAssets(fa || []);
      setLoadError("");
    } catch (e) {
      setLoadError(e.message || "Ma'lumotlarni yuklashda xatolik. .env sozlamalarini tekshiring.");
    } finally {
      setReady(true);
    }
  };

  useEffect(() => { reloadAll(); }, []);

  const balances = useMemo(() => {
    const map = {};
    customers.forEach((c) => { map[c.id] = { invoiced: 0, paid: 0 }; });
    invoices.forEach((i) => { if (map[i.customer_id]) map[i.customer_id].invoiced += Number(i.total); });
    payments.forEach((p) => { if (map[p.customer_id]) map[p.customer_id].paid += Number(p.amount); });
    return map;
  }, [customers, invoices, payments]);

  const stockValue = products.reduce((s, p) => s + Number(p.qty) * Number(p.cost_price > 0 ? p.cost_price : p.price), 0);
  const totalDebt = Object.values(balances).reduce((s, b) => s + Math.max(b.invoiced - b.paid, 0), 0);
  const totalCredit = Object.values(balances).reduce((s, b) => s + Math.max(b.paid - b.invoiced, 0), 0);

  const rawMaterialsValue = rawMaterialBatches.filter((b) => b.active).reduce((s, b) => s + Number(b.remaining_qty) * Number(b.unit_cost), 0);
  const frozenMaterialsValue = rawMaterialBatches.filter((b) => !b.active).reduce((s, b) => s + Number(b.remaining_qty) * Number(b.unit_cost), 0);
  const cashBalanceSom = cashTransactions.filter((t) => (t.currency || "som") === "som").reduce((s, t) => s + (t.type === "kirim" ? Number(t.amount) : -Number(t.amount)), 0);
  const cashBalanceUsd = cashTransactions.filter((t) => t.currency === "usd").reduce((s, t) => s + (t.type === "kirim" ? Number(t.amount) : -Number(t.amount)), 0);

  const exportExcel = () => exportAllToExcel({
    products, customers, invoices, payments, rawMaterials, rawMaterialBatches, finishedProducts, productionBatches, cashTransactions,
    balances, rawMaterialsValue, stockValue, totalDebt, totalCredit, cashBalanceSom, cashBalanceUsd, usdRate: Number(settings.usd_rate) || 0,
  });

  if (!ready) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: "#5B6660" }}>
        Yuklanmoqda…
      </div>
    );
  }

  if (settings.app_password && !unlocked) {
    const tryUnlock = () => {
      if (passwordInput === settings.app_password) {
        setUnlocked(true);
        setPasswordError("");
        try { sessionStorage.setItem("ombor_unlocked", "1"); } catch (e) {}
      } else {
        setPasswordError("Parol noto'g'ri");
      }
    };
    return (
      <div className="app-shell">
        <style>{CSS}</style>
        <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div className="modal-card" style={{ maxWidth: 340 }}>
            <div className="card-title">{settings.company_name || "Ombor Hisobi"}</div>
            <div className="muted" style={{ marginBottom: 10, fontSize: 13 }}>Kirish uchun parolni kiriting</div>
            <input
              className="input" type="password" placeholder="Parol" value={passwordInput}
              onChange={(e) => { setPasswordInput(e.target.value); setPasswordError(""); }}
              onKeyDown={(e) => { if (e.key === "Enter") tryUnlock(); }}
            />
            {passwordError && <div className="error-box" style={{ marginTop: 10 }}><AlertTriangle size={14} /> {passwordError}</div>}
            <button className="btn btn-primary" style={{ marginTop: 14, width: "100%", justifyContent: "center" }} onClick={tryUnlock}>Kirish</button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <style>{CSS}</style>

      <header className="topbar">
        <div className="brand" role="button" tabIndex={0} onClick={() => setShowCompanyModal(true)} style={{ cursor: "pointer" }}>
          <span className="brand-mark">§</span>
          <div>
            <div className="brand-title">{settings.company_name || "Ombor Hisobi"}</div>
            <div className="brand-sub">Ombor · Faktura · Hisob-kitob</div>
          </div>
        </div>
        <button className="btn btn-ghost" style={{ marginLeft: "auto" }} onClick={() => setShowUserModal(true)}>
          <Users size={14} /> {currentUser || "Foydalanuvchi tanlash"}
        </button>
      </header>

      {loadError && (
        <div className="error-box" style={{ margin: "0 20px" }}>
          <AlertTriangle size={14} /> {loadError}
        </div>
      )}

      <nav className="tabs">
        {[
          { id: "dashboard", label: "Bosh sahifa", icon: LayoutDashboard },
          { id: "products", label: "Ombor", icon: Package },
          { id: "production", label: "Ishlab chiqarish", icon: Factory },
          { id: "invoice", label: "Faktura", icon: FileText },
          { id: "kassa", label: "Kassa", icon: Banknote },
          { id: "ledger", label: "Hisob-kitob", icon: BookOpen },
          { id: "customers", label: "Mijozlar", icon: Users },
          { id: "suppliers", label: "Ta'minotchilar", icon: Truck },
          { id: "analytics", label: "Analiz", icon: BarChart3 },
        ].map(({ id, label, icon: Icon }) => (
          <button key={id} className={`tab ${tab === id ? "active" : ""}`} onClick={() => setTab(id)}>
            <Icon size={16} strokeWidth={2} />
            <span>{label}</span>
          </button>
        ))}
      </nav>

      <main className="main">
        {tab === "dashboard" && (
          <Dashboard
            rawMaterialsValue={rawMaterialsValue} frozenMaterialsValue={frozenMaterialsValue} stockValue={stockValue} totalDebt={totalDebt} totalCredit={totalCredit}
            cashBalanceSom={cashBalanceSom} cashBalanceUsd={cashBalanceUsd} usdRate={Number(settings.usd_rate) || 0}
            invoices={invoices} onExport={exportExcel}
            products={products} customers={customers} payments={payments} movements={movements}
            rawMaterialBatches={rawMaterialBatches} productionConsumptions={productionConsumptions} productionBatches={productionBatches}
            cashTransactions={cashTransactions} rawMaterials={rawMaterials} settings={settings} onReload={reloadAll}
            fixedAssets={fixedAssets}
          />
        )}
        {tab === "products" && (
          <Products
            products={products} movements={movements} rawMaterials={rawMaterials} rawMaterialBatches={rawMaterialBatches}
            productionBatches={productionBatches} productionConsumptions={productionConsumptions} suppliers={suppliers} onReload={reloadAll}
          />
        )}
        {tab === "production" && (
          <Production
            products={products} finishedProducts={finishedProducts} rawMaterials={rawMaterials}
            rawMaterialBatches={rawMaterialBatches} productNorms={productNorms} productionBatches={productionBatches}
            productionConsumptions={productionConsumptions} cashTransactions={cashTransactions} currentUser={currentUser}
            onReload={reloadAll} prefill={productionPrefill} onPrefillConsumed={() => setProductionPrefill(null)}
            fixedAssets={fixedAssets}
          />
        )}
        {tab === "customers" && <Customers customers={customers} onReload={reloadAll} />}
        {tab === "suppliers" && (
          <Suppliers suppliers={suppliers} supplierPayments={supplierPayments} rawMaterialBatches={rawMaterialBatches} rawMaterials={rawMaterials} onReload={reloadAll} />
        )}
        {tab === "invoice" && (
          <InvoiceTab
            products={products} customers={customers} invoices={invoices} currentUser={currentUser}
            onReload={reloadAll} onPrint={setPrintInvoice}
            onNeedProduction={(productId, shortfall) => { setProductionPrefill({ productId, qty: shortfall }); setTab("production"); }}
            customerId={invCustomerId} setCustomerId={setInvCustomerId}
            date={invDate} setDate={setInvDate}
            rows={invRows} setRows={setInvRows}
          />
        )}
        {tab === "kassa" && (
          <Kassa cashTransactions={cashTransactions} cashCategories={cashCategories} customers={customers} suppliers={suppliers} settings={settings} currentUser={currentUser} onReload={reloadAll} />
        )}
        {tab === "ledger" && (
          <Ledger customers={customers} invoices={invoices} payments={payments} balances={balances} settings={settings} onReload={reloadAll} />
        )}
        {tab === "analytics" && (
          <Analytics
            products={products} customers={customers} invoices={invoices} payments={payments} cashTransactions={cashTransactions} balances={balances}
            finishedProducts={finishedProducts} productNorms={productNorms} rawMaterials={rawMaterials} rawMaterialBatches={rawMaterialBatches}
            productionBatches={productionBatches} fixedAssets={fixedAssets}
          />
        )}
      </main>

      {printInvoice && (
        <PrintOverlay invoice={printInvoice} customers={customers} settings={settings} onClose={() => setPrintInvoice(null)} />
      )}
      {showCompanyModal && (
        <CompanyModal settings={settings} onClose={() => setShowCompanyModal(false)} onSaved={reloadAll} />
      )}
      {showUserModal && (
        <UserModal currentUser={currentUser} onClose={() => setShowUserModal(false)} onSelect={(name) => { setCurrentUser(name); try { localStorage.setItem("ombor_user", name); } catch (e) {} setShowUserModal(false); }} />
      )}
    </div>
  );
}

function UserModal({ currentUser, onClose, onSelect }) {
  const [custom, setCustom] = useState("");
  const presets = ["Foydalanuvchi 1", "Foydalanuvchi 2", "Foydalanuvchi 3", "Foydalanuvchi 4", "Foydalanuvchi 5"];
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <div className="card-title">Siz kimsiz?</div>
        <div className="muted" style={{ marginBottom: 10, fontSize: 12.5 }}>
          Bu — parol emas, faqat kim nima kiritganini bilish uchun. Ismingizni tanlang yoki yozing.
        </div>
        <div className="form-row wrap" style={{ marginBottom: 10 }}>
          {presets.map((p) => (
            <button key={p} className={`btn ${currentUser === p ? "btn-primary" : "btn-ghost"}`} onClick={() => onSelect(p)}>{p}</button>
          ))}
        </div>
        <div className="form-row wrap">
          <input className="input" placeholder="Yoki ismingizni yozing" value={custom} onChange={(e) => setCustom(e.target.value)} />
          <button className="btn btn-primary" onClick={() => custom.trim() && onSelect(custom.trim())}><Check size={14} /> Tanlash</button>
        </div>
      </div>
    </div>
  );
}

function CompanyModal({ settings, onClose, onSaved }) {
  const [name, setName] = useState(settings.company_name);
  const [phone, setPhone] = useState(settings.company_phone);
  const [appPassword, setAppPassword] = useState(settings.app_password || "");
  const [tgToken, setTgToken] = useState(settings.telegram_bot_token || "");
  const [tgChat, setTgChat] = useState(settings.telegram_chat_id || "");
  const [saving, setSaving] = useState(false);

  const save = async () => {
    setSaving(true);
    try {
      await api.settings.save({
        ...settings, company_name: name, company_phone: phone,
        app_password: appPassword, telegram_bot_token: tgToken, telegram_chat_id: tgChat,
      });
      await onSaved(); onClose();
    } finally { setSaving(false); }
  };

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-card" onClick={(e) => e.stopPropagation()}>
        <div className="card-title">Korxona ma'lumotlari</div>
        <div className="form-row wrap" style={{ marginBottom: 6 }}>
          <input className="input" placeholder="Korxona nomi" value={name} onChange={(e) => setName(e.target.value)} />
        </div>
        <div className="form-row wrap" style={{ marginBottom: 14 }}>
          <input className="input" placeholder="Telefon" value={phone} onChange={(e) => setPhone(e.target.value)} />
        </div>

        <div className="card-title" style={{ fontSize: 13 }}>Xavfsizlik</div>
        <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>Saytga kirish uchun umumiy parol (bo'sh qoldirsangiz — parolsiz, hammaga ochiq)</div>
        <div className="form-row wrap" style={{ marginBottom: 14 }}>
          <input className="input" placeholder="Parol (ixtiyoriy)" value={appPassword} onChange={(e) => setAppPassword(e.target.value)} />
        </div>

        <div className="card-title" style={{ fontSize: 13 }}>Telegram xabarnoma</div>
        <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>Bosh sahifadagi "Telegram'ga yuborish" tugmasi shu bot orqali xabar yuboradi</div>
        <div className="form-row wrap" style={{ marginBottom: 6 }}>
          <input className="input" placeholder="Bot token" value={tgToken} onChange={(e) => setTgToken(e.target.value)} />
        </div>
        <div className="form-row wrap">
          <input className="input" placeholder="Chat ID" value={tgChat} onChange={(e) => setTgChat(e.target.value)} />
        </div>

        <div className="invoice-footer" style={{ marginTop: 14 }}>
          <button className="btn btn-ghost" onClick={onClose}>Yopish</button>
          <button className="btn btn-primary" disabled={saving} onClick={save}><Check size={14} /> Saqlash</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------- Dashboard ---------------- */
function computeBalanceAsOf(date, data) {
  const { products, customers, invoices, payments, movements, rawMaterialBatches, productionConsumptions, productionBatches, cashTransactions, usdRate } = data;

  const batchDateMap = {};
  productionBatches.forEach((pb) => { batchDateMap[pb.id] = pb.date; });

  let rawMaterialsValue = 0;
  rawMaterialBatches.forEach((b) => {
    if (b.date > date) return;
    const consumedByThen = productionConsumptions
      .filter((c) => c.raw_material_batch_id === b.id && batchDateMap[c.production_batch_id] && batchDateMap[c.production_batch_id] <= date)
      .reduce((s, c) => s + Number(c.qty), 0);
    const remaining = Number(b.qty) - consumedByThen;
    rawMaterialsValue += remaining * Number(b.unit_cost);
  });

  let stockValue = 0;
  products.forEach((p) => {
    const qtyAsOf = movements.filter((m) => m.product_id === p.id && m.date <= date).reduce((s, m) => s + Number(m.qty), 0);
    stockValue += qtyAsOf * Number(p.cost_price > 0 ? p.cost_price : p.price);
  });

  let totalDebt = 0, totalCredit = 0;
  customers.forEach((c) => {
    const invoicedAsOf = invoices.filter((i) => i.customer_id === c.id && i.date <= date).reduce((s, i) => s + Number(i.total), 0);
    const paidAsOf = payments.filter((p) => p.customer_id === c.id && p.date <= date).reduce((s, p) => s + Number(p.amount), 0);
    const bal = invoicedAsOf - paidAsOf;
    if (bal > 0) totalDebt += bal; else totalCredit += -bal;
  });

  const cashBalanceSom = cashTransactions.filter((t) => (t.currency || "som") === "som" && t.date <= date).reduce((s, t) => s + (t.type === "kirim" ? Number(t.amount) : -Number(t.amount)), 0);
  const cashBalanceUsd = cashTransactions.filter((t) => t.currency === "usd" && t.date <= date).reduce((s, t) => s + (t.type === "kirim" ? Number(t.amount) : -Number(t.amount)), 0);

  return { rawMaterialsValue, stockValue, totalDebt, totalCredit, cashBalanceSom, cashBalanceUsd };
}

function Dashboard({
  rawMaterialsValue, frozenMaterialsValue, stockValue, totalDebt, totalCredit, cashBalanceSom, cashBalanceUsd, usdRate, invoices, onExport,
  products, customers, payments, movements, rawMaterialBatches, productionConsumptions, productionBatches, cashTransactions, rawMaterials, settings, onReload,
  fixedAssets,
}) {
  const [asOfDate, setAsOfDate] = useState("");
  const [sending, setSending] = useState(false);
  const [sendMsg, setSendMsg] = useState("");

  const shown = useMemo(() => {
    if (!asOfDate) return { rawMaterialsValue, stockValue, totalDebt, totalCredit, cashBalanceSom, cashBalanceUsd };
    return computeBalanceAsOf(asOfDate, {
      products, customers, invoices, payments, movements, rawMaterialBatches, productionConsumptions, productionBatches, cashTransactions, usdRate,
    });
  }, [asOfDate, rawMaterialsValue, stockValue, totalDebt, totalCredit, cashBalanceSom, cashBalanceUsd]);

  const cashInSom = shown.cashBalanceSom + shown.cashBalanceUsd * usdRate;
  const jamiAktivlar = shown.rawMaterialsValue + shown.stockValue + shown.totalDebt + cashInSom;
  const jamiPassivlar = shown.totalCredit;
  const qoldiq = jamiAktivlar - jamiPassivlar;

  const lowStockMaterials = (rawMaterials || []).map((m) => {
    const activeQty = (rawMaterialBatches || []).filter((b) => b.raw_material_id === m.id && b.active).reduce((s, b) => s + Number(b.remaining_qty), 0);
    return { material: m, activeQty };
  }).filter((r) => Number(r.material.reorder_point) > 0 && r.activeQty <= Number(r.material.reorder_point));

  const sendTelegram = async () => {
    if (!settings.telegram_bot_token || !settings.telegram_chat_id) {
      setSendMsg("Avval yuqoridagi logotipni bosib, Telegram bot token va Chat ID kiriting");
      return;
    }
    setSending(true);
    setSendMsg("");
    try {
      let text = `📊 ${settings.company_name || "Ombor Hisobi"} — balans holati\n\n`;
      text += `Xom ashyo ombori: ${fmt(rawMaterialsValue)}\nTayyor mahsulot ombori: ${fmt(stockValue)}\nDebitor: ${fmt(totalDebt)}\nKreditor: ${fmt(totalCredit)}\nKassa (so'm+dollar): ${fmt(cashBalanceSom + cashBalanceUsd * usdRate)}\n`;
      if (lowStockMaterials.length > 0) {
        text += `\n⚠️ Kam qolgan xom ashyo:\n`;
        lowStockMaterials.forEach((r) => { text += `— ${r.material.name}: ${fmtQty(r.activeQty)} ${r.material.unit} (chegara: ${r.material.reorder_point})\n`; });
      }
      const url = `https://api.telegram.org/bot${settings.telegram_bot_token}/sendMessage`;
      const res = await fetch(url, {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ chat_id: settings.telegram_chat_id, text }),
      });
      const data = await res.json();
      if (!data.ok) throw new Error(data.description || "Telegram xatosi");
      await api.settings.save({ ...settings, last_telegram_sent: todayISO() });
      setSendMsg("Xabar muvaffaqiyatli yuborildi ✅");
      if (onReload) await onReload();
    } catch (e) {
      setSendMsg("Xatolik: " + (e.message || "yuborib bo'lmadi"));
    } finally {
      setSending(false);
    }
  };

  return (
    <div>
      <div className="card">
        <div className="card-title-row">
          <div className="card-title" style={{ fontSize: 18 }}>Balans {asOfDate && <span className="muted" style={{ fontSize: 13, fontWeight: 400 }}>— {fmtDate(asOfDate)} holatiga</span>}</div>
          <button className="btn btn-primary" onClick={onExport}><FileText size={14} /> Excel yuklab olish</button>
        </div>

        <div className="form-row wrap" style={{ marginBottom: 14 }}>
          <label className="muted" style={{ fontSize: 12.5 }}>Qaysi sanaga balansni ko'rish:</label>
          <input className="input xs" type="date" value={asOfDate} onChange={(e) => setAsOfDate(e.target.value)} />
          {asOfDate && <button className="btn btn-ghost" onClick={() => setAsOfDate("")}>Joriy holatga qaytish</button>}
        </div>

        <div className="balance-columns">
          <div>
            <div className="balance-col-title">AKTIVLAR</div>
            <table className="table">
              <tbody>
                <tr><td>Xom ashyo ombori{asOfDate ? "" : " (faol)"}</td><td className="right mono">{fmt(shown.rawMaterialsValue)}</td></tr>
                {!asOfDate && frozenMaterialsValue > 0 && (
                  <tr><td className="muted">— shu jumladan harakatsiz, kirmaydi</td><td className="right mono muted">{fmt(frozenMaterialsValue)}</td></tr>
                )}
                <tr><td>Tayyor mahsulot ombori</td><td className="right mono">{fmt(shown.stockValue)}</td></tr>
                <tr><td>Debitor (mijozlar qarzi)</td><td className="right mono">{fmt(shown.totalDebt)}</td></tr>
                <tr>
                  <td>Kassa va hisob raqami</td>
                  <td className="right mono">
                    {fmt(cashInSom)}
                    <div className="muted" style={{ fontSize: 11 }}>
                      {fmt(shown.cashBalanceSom)} + ${fmtNum(shown.cashBalanceUsd)} × {usdRate || 0}
                    </div>
                  </td>
                </tr>
                <tr className="editing-row"><td><b>Jami aktivlar</b></td><td className="right mono"><b>{fmt(jamiAktivlar)}</b></td></tr>
              </tbody>
            </table>
          </div>

          <div>
            <div className="balance-col-title">PASSIVLAR</div>
            <table className="table">
              <tbody>
                <tr><td>Kreditor (mijozlarga ortiqcha to'langan)</td><td className="right mono">{fmt(shown.totalCredit)}</td></tr>
                <tr className="editing-row"><td><b>Jami passivlar</b></td><td className="right mono"><b>{fmt(jamiPassivlar)}</b></td></tr>
              </tbody>
            </table>
          </div>
        </div>

        <div className="balance-net">
          <span>Qoldiq (Jami aktiv − Jami passiv):</span>
          <span className={`mono balance-net-amount ${qoldiq < 0 ? "tone-debt-text" : "tone-ok-text"}`}>{fmt(qoldiq)}</span>
        </div>
      </div>

      <div className="card">
        <div className="card-title">Statistika</div>
        <div className="muted">Jami chiqarilgan fakturalar: <b>{invoices.length}</b> ta</div>
      </div>

      <FixedAssetsCard fixedAssets={fixedAssets} onReload={onReload} />
    </div>
  );
}

function FixedAssetsCard({ fixedAssets, onReload }) {
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ name: "", purchase_price: "", purchase_date: todayISO(), depreciation_rate: "", include_in_cost: true, note: "" });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [busy, setBusy] = useState(false);

  const monthlyDep = (a) => (Number(a.purchase_price) || 0) * (Number(a.depreciation_rate) || 0) / 100 / 12;

  const add = async () => {
    if (!form.name.trim()) return;
    setBusy(true);
    try {
      await api.fixedAssets.add({
        name: form.name.trim(), purchase_price: Number(form.purchase_price) || 0, purchase_date: form.purchase_date || null,
        depreciation_rate: Number(form.depreciation_rate) || 0, include_in_cost: form.include_in_cost, note: form.note,
      });
      setForm({ name: "", purchase_price: "", purchase_date: todayISO(), depreciation_rate: "", include_in_cost: true, note: "" });
      setShowAdd(false);
      await onReload();
    } finally { setBusy(false); }
  };
  const startEdit = (a) => { setEditingId(a.id); setEditForm({ ...a }); };
  const saveEdit = async () => {
    setBusy(true);
    try {
      await api.fixedAssets.update(editingId, {
        name: editForm.name, purchase_price: Number(editForm.purchase_price) || 0, purchase_date: editForm.purchase_date || null,
        depreciation_rate: Number(editForm.depreciation_rate) || 0, include_in_cost: !!editForm.include_in_cost, note: editForm.note,
      });
      setEditingId(null);
      await onReload();
    } finally { setBusy(false); }
  };
  const remove = async (id) => {
    if (!confirm("Bu asosiy vositani o'chirasizmi?")) return;
    setBusy(true);
    try { await api.fixedAssets.remove(id); await onReload(); } finally { setBusy(false); }
  };

  const totalValue = fixedAssets.reduce((s, a) => s + (Number(a.purchase_price) || 0), 0);
  const totalMonthlyDep = fixedAssets.filter((a) => a.include_in_cost).reduce((s, a) => s + monthlyDep(a), 0);

  return (
    <div className="card">
      <div className="card-title-row">
        <div className="card-title">Asosiy vositalar</div>
        <button className="btn btn-primary" onClick={() => setShowAdd(true)}><Plus size={14} /> Yangi vosita</button>
      </div>
      <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>
        Uskuna, bino, transport va h.k. Amortizatsiya foizi — yillik. "Tannarxga kirsin" belgilangan vositalarning oylik amortizatsiyasi ishlab chiqarish tannarxidagi operatsion xarajat ulushiga qo'shiladi.
      </div>
      <table className="table">
        <thead><tr><th>Nomi</th><th className="right">Narxi</th><th>Sotib olingan sana</th><th className="right">Amortizatsiya %/yil</th><th className="right">Oylik amortizatsiya</th><th>Tannarxga kirsinmi</th><th></th></tr></thead>
        <tbody>
          {fixedAssets.map((a) => editingId === a.id ? (
            <tr key={a.id} className="editing-row">
              <td><input className="input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} /></td>
              <td><input className="input mono" type="number" value={editForm.purchase_price} onChange={(e) => setEditForm({ ...editForm, purchase_price: e.target.value })} /></td>
              <td><input className="input xs" type="date" value={editForm.purchase_date || ""} onChange={(e) => setEditForm({ ...editForm, purchase_date: e.target.value })} /></td>
              <td><input className="input mono xs" type="number" value={editForm.depreciation_rate} onChange={(e) => setEditForm({ ...editForm, depreciation_rate: e.target.value })} /></td>
              <td className="right mono">{fmt(monthlyDep(editForm))}</td>
              <td><input type="checkbox" checked={!!editForm.include_in_cost} onChange={(e) => setEditForm({ ...editForm, include_in_cost: e.target.checked })} /></td>
              <td className="row-actions"><button className="icon-btn" disabled={busy} onClick={saveEdit}><Check size={15} /></button></td>
            </tr>
          ) : (
            <tr key={a.id}>
              <td>{a.name}{a.note && <div className="muted" style={{ fontSize: 11 }}>{a.note}</div>}</td>
              <td className="right mono">{fmt(a.purchase_price)}</td>
              <td className="muted">{a.purchase_date ? fmtDate(a.purchase_date) : "—"}</td>
              <td className="right mono">{a.depreciation_rate}%</td>
              <td className="right mono">{fmt(monthlyDep(a))}</td>
              <td>{a.include_in_cost ? <span className="tone-ok-text">Ha</span> : <span className="muted">Yo'q</span>}</td>
              <td className="row-actions">
                <button className="icon-btn" onClick={() => startEdit(a)}><Pencil size={14} /></button>
                <button className="icon-btn danger" onClick={() => remove(a.id)}><Trash2 size={14} /></button>
              </td>
            </tr>
          ))}
          {fixedAssets.length === 0 && <tr><td colSpan={7} className="empty">Hali asosiy vosita qo'shilmagan</td></tr>}
        </tbody>
        {fixedAssets.length > 0 && (
          <tfoot>
            <tr className="editing-row">
              <td><b>Jami</b></td>
              <td className="right mono"><b>{fmt(totalValue)}</b></td>
              <td colSpan={2}></td>
              <td className="right mono"><b>{fmt(totalMonthlyDep)}</b></td>
              <td colSpan={2} className="muted" style={{ fontSize: 11 }}>tannarxga kiruvchi oylik jami</td>
            </tr>
          </tfoot>
        )}
      </table>

      {showAdd && (
        <div className="modal-backdrop" onClick={() => setShowAdd(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="card-title">Yangi asosiy vosita</div>
            <div className="form-row wrap" style={{ marginBottom: 8 }}>
              <input className="input" placeholder="Nomi (masalan: Idishlash liniyasi)" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </div>
            <div className="form-row wrap" style={{ marginBottom: 8 }}>
              <input className="input sm mono" type="number" placeholder="Narxi" value={form.purchase_price} onChange={(e) => setForm({ ...form, purchase_price: e.target.value })} />
              <input className="input xs" type="date" value={form.purchase_date} onChange={(e) => setForm({ ...form, purchase_date: e.target.value })} />
              <input className="input xs mono" type="number" placeholder="Amortizatsiya %/yil" value={form.depreciation_rate} onChange={(e) => setForm({ ...form, depreciation_rate: e.target.value })} />
            </div>
            <div className="form-row wrap" style={{ marginBottom: 8 }}>
              <input className="input" placeholder="Izoh (ixtiyoriy)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
            </div>
            <label className="form-row" style={{ gap: 6, cursor: "pointer" }}>
              <input type="checkbox" checked={form.include_in_cost} onChange={(e) => setForm({ ...form, include_in_cost: e.target.checked })} />
              <span className="muted" style={{ fontSize: 13 }}>Oylik amortizatsiyasi ishlab chiqarish tannarxiga qo'shilsin</span>
            </label>
            <div className="invoice-footer" style={{ marginTop: 14 }}>
              <button className="btn btn-ghost" onClick={() => setShowAdd(false)}>Bekor qilish</button>
              <button className="btn btn-primary" disabled={busy} onClick={add}><Plus size={14} /> Qo'shish</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function StatCard({ label, value, mono, tone }) {
  return (
    <div className={`stat-card ${tone ? "tone-" + tone : ""}`}>
      <div className="stat-label">{label}</div>
      <div className={`stat-value ${mono ? "mono" : ""}`}>{value}</div>
    </div>
  );
}

/* ---------------- Kassa ---------------- */
const CASH_CATEGORIES = [
  "Sotuvdan tushum",
  "Boshqa tushum",
  "Kapital kiritish",
  "Xom ashyo xaridi",
  "Ijara",
  "Ish haqi",
  "Kommunal xizmatlar",
  "Transport",
  "Soliq",
  "Ta'mirlash",
  "Reklama",
  "Bank xizmati",
  "Aloqa (telefon/internet)",
  "Ofis xarajatlari",
  "Boshqa xarajat",
];

function Kassa({ cashTransactions, cashCategories, customers, suppliers, settings, currentUser, onReload }) {
  const [currency, setCurrency] = useState("som");
  const allCategories = useMemo(() => {
    const custom = cashCategories.map((c) => c.name);
    return [...CASH_CATEGORIES, ...custom.filter((n) => !CASH_CATEGORIES.includes(n))];
  }, [cashCategories]);
  const [form, setForm] = useState({ date: todayISO(), note: "", category: CASH_CATEGORIES[0], type: "kirim", amount: "", customerId: "", supplierId: "", rate: "" });
  const [usdRate, setUsdRate] = useState(settings.usd_rate || "");
  const [savingRate, setSavingRate] = useState(false);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [showNewCategory, setShowNewCategory] = useState(false);
  const [newCategoryName, setNewCategoryName] = useState("");
  const [showSettings, setShowSettings] = useState(false);

  const addCategory = async () => {
    const name = newCategoryName.trim();
    if (!name) return;
    if (allCategories.includes(name)) { setForm({ ...form, category: name }); setShowNewCategory(false); setNewCategoryName(""); return; }
    try {
      await api.cashCategories.add({ name });
      setForm({ ...form, category: name });
      setShowNewCategory(false); setNewCategoryName("");
      await onReload();
    } catch (e) {
      setError("Bo'lim qo'shishda xatolik: " + (e.message || "noma'lum xato"));
    }
  };

  const isSaleIncome = form.category === "Sotuvdan tushum" && form.type === "kirim";
  const isSupplierPayment = form.category === "Xom ashyo xaridi" && form.type === "chiqim";
  const isUsd = currency === "usd";

  const addTx = async () => {
    if (busy) return;
    setError("");
    if (!Number(form.amount)) return;
    if (isSaleIncome && !form.customerId) { setError("Sotuvdan tushum uchun mijozni tanlang"); return; }
    if (isSupplierPayment && !form.supplierId) { setError("Ta'minotchiga to'lov uchun ta'minotchini tanlang"); return; }
    if (isUsd && !Number(form.rate)) { setError("Dollar uchun kursni kiriting"); return; }
    setBusy(true);
    try {
      const rate = isUsd ? Number(form.rate) : null;
      let note = form.note;
      if (isSaleIncome) note = `${customers.find((c) => c.id === form.customerId)?.name || ""}${form.note ? " — " + form.note : ""}`;
      if (isSupplierPayment) note = `${suppliers.find((s) => s.id === form.supplierId)?.name || ""}${form.note ? " — " + form.note : ""}`;
      let paymentId = null;
      let supplierPaymentId = null;
      const somAmount = isUsd ? Number(form.amount) * rate : Number(form.amount);
      if (isSaleIncome) {
        const payment = await api.payments.add({
          customer_id: form.customerId, amount: somAmount, date: form.date,
          note: isUsd ? `${form.note ? form.note + " — " : ""}$${form.amount} × ${rate}` : form.note,
          created_by: currentUser || null,
        });
        paymentId = payment.id;
      }
      if (isSupplierPayment) {
        const supplierPayment = await api.supplierPayments.add({
          supplier_id: form.supplierId, amount: somAmount, date: form.date,
          note: isUsd ? `${form.note ? form.note + " — " : ""}$${form.amount} × ${rate}` : form.note,
        });
        supplierPaymentId = supplierPayment.id;
      }
      await api.cashTransactions.add({
        date: form.date, note, category: form.category, type: form.type, amount: Number(form.amount), currency,
        payment_id: paymentId, supplier_payment_id: supplierPaymentId, rate,
        created_by: currentUser || null,
      });
      setForm({ date: todayISO(), note: "", category: CASH_CATEGORIES[0], type: "kirim", amount: "", customerId: "", supplierId: "", rate: "" });
      await onReload();
    } catch (e) {
      setError("Saqlashda xatolik: " + (e.message || "noma'lum xato") + " — Supabase'da so'nggi SQL fayllar ishga tushirilganini tekshiring.");
    } finally {
      setBusy(false);
    }
  };
  const [listError, setListError] = useState("");
  const removeTx = async (t) => {
    if (!confirm((t.payment_id || t.supplier_payment_id) ? "Bu yozuv o'chirilsa, bog'liq to'lov ham Hisob-kitobdan o'chadi. Davom etasizmi?" : "Bu yozuvni o'chirasizmi?")) return;
    try {
      if (t.payment_id) await api.payments.remove(t.payment_id);
      if (t.supplier_payment_id) await api.supplierPayments.remove(t.supplier_payment_id);
      await api.cashTransactions.remove(t.id);
      await onReload();
    } catch (e) {
      setListError("O'chirishda xatolik: " + (e.message || "noma'lum xato"));
    }
  };

  const saveRate = async () => {
    setSavingRate(true);
    try { await api.settings.save({ ...settings, usd_rate: Number(usdRate) || 0 }); await onReload(); }
    finally { setSavingRate(false); }
  };

  const list = cashTransactions
    .filter((t) => (t.currency || "som") === currency)
    .sort((a, b) => new Date(a.date) - new Date(b.date) || new Date(a.created_at) - new Date(b.created_at));

  let running = 0;
  const rows = list.map((t) => {
    running += t.type === "kirim" ? Number(t.amount) : -Number(t.amount);
    return { ...t, balance: running };
  });

  const symbol = currency === "usd" ? "$" : "so'm";

  const [asOfDate, setAsOfDate] = useState("");
  const asOfBalance = asOfDate
    ? rows.filter((t) => t.date <= asOfDate).reduce((s, t) => s + (t.type === "kirim" ? Number(t.amount) : -Number(t.amount)), 0)
    : null;

  return (
    <div>
      {listError && <div className="error-box" style={{ marginBottom: 12 }}><AlertTriangle size={14} /> {listError}</div>}

      <div className="card-title-row" style={{ marginBottom: 8 }}>
        <div className="tabs" style={{ padding: 0, border: "none", marginBottom: 0 }}>
          <button className={`tab ${currency === "som" ? "active" : ""}`} onClick={() => setCurrency("som")}>So'm kassa</button>
          <button className={`tab ${currency === "usd" ? "active" : ""}`} onClick={() => { setCurrency("usd"); setForm((f) => ({ ...f, rate: f.rate || usdRate })); }}>Dollar kassa</button>
        </div>
        <button className="btn btn-ghost" onClick={() => setShowSettings(true)}><Settings2 size={14} /> Sozlamalar</button>
      </div>

      <div className="card">
        <div className="card-title">Yangi yozuv — {currency === "usd" ? "Dollar" : "So'm"} kassa</div>
        <div className="form-row wrap">
          <input className="input xs" type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
          <input className="input" placeholder="Mazmuni (izoh)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
          <select className="input" value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>
            {allCategories.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <button className="btn btn-ghost" style={{ padding: "8px 10px" }} onClick={() => setShowNewCategory(true)}><Plus size={13} /> Yangi bo'lim</button>
          <select className="input xs" value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
            <option value="kirim">Kirim</option>
            <option value="chiqim">Chiqim</option>
          </select>
          <input className="input sm mono" type="number" placeholder="Summa ($)" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} />
          {isUsd && (
            <input className="input sm mono" type="number" placeholder="Kurs (1$=...)" value={form.rate} onChange={(e) => setForm({ ...form, rate: e.target.value })} />
          )}
        </div>
        {isUsd && Number(form.amount) > 0 && Number(form.rate) > 0 && (
          <div className="muted" style={{ fontSize: 12.5, marginTop: 6 }}>≈ {fmt(Number(form.amount) * Number(form.rate))} (shu kursda)</div>
        )}
        {isSaleIncome && (
          <div className="form-row wrap" style={{ marginTop: 8 }}>
            <select className="input" value={form.customerId} onChange={(e) => setForm({ ...form, customerId: e.target.value })}>
              <option value="">— Mijozni tanlang —</option>
              {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
        )}
        {isSupplierPayment && (
          <div className="form-row wrap" style={{ marginTop: 8 }}>
            <select className="input" value={form.supplierId} onChange={(e) => setForm({ ...form, supplierId: e.target.value })}>
              <option value="">— Ta'minotchini tanlang —</option>
              {suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </div>
        )}
        {error && <div className="error-box"><AlertTriangle size={14} /> {error}</div>}
        <div className="invoice-footer" style={{ marginTop: 10 }}>
          <div />
          <button className="btn btn-primary" disabled={busy} onClick={addTx}><Plus size={14} /> {busy ? "Saqlanmoqda…" : "Qo'shish"}</button>
        </div>
      </div>

      <div className="card">
        <div className="card-title-row">
          <div className="card-title">{currency === "usd" ? "Dollar" : "So'm"} kassa harakatlari ({rows.length})</div>
          <div className="mono" style={{ fontWeight: 700 }}>Qoldiq: {fmtNum(running)} {symbol}</div>
        </div>
        <table className="table">
          <thead><tr><th>№</th><th>Sana</th><th>Mazmuni</th><th>Bo'lim</th>{isUsd && <th className="right">Kurs</th>}<th className="right">Kirim</th><th className="right">Chiqim</th><th className="right">Qoldiq</th><th>Kim</th><th></th></tr></thead>
          <tbody>
            {rows.map((t, idx) => (
              <tr key={t.id}>
                <td className="muted">{idx + 1}</td>
                <td className="muted">{fmtDate(t.date)}</td>
                <td>{t.note || "—"}</td>
                <td className="muted">{t.category || "—"}</td>
                {isUsd && <td className="right mono muted">{t.rate || "—"}</td>}
                <td className="right mono tone-ok-text">{t.type === "kirim" ? fmtNum(t.amount) : "—"}</td>
                <td className="right mono tone-debt-text">{t.type === "chiqim" ? fmtNum(t.amount) : "—"}</td>
                <td className="right mono">{fmtNum(t.balance)}</td>
                <td className="muted">{t.created_by || "—"}</td>
                <td className="row-actions"><button className="icon-btn danger" onClick={() => removeTx(t)}><Trash2 size={14} /></button></td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={isUsd ? 10 : 9} className="empty">Hali yozuv yo'q</td></tr>}
          </tbody>
        </table>
      </div>

      {showNewCategory && (
        <div className="modal-backdrop" onClick={() => setShowNewCategory(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="card-title">Yangi bo'lim (kategoriya) qo'shish</div>
            <div className="form-row wrap">
              <input className="input" placeholder="Masalan: Transport xarajati" value={newCategoryName} onChange={(e) => setNewCategoryName(e.target.value)} />
            </div>
            <div className="invoice-footer" style={{ marginTop: 14 }}>
              <button className="btn btn-ghost" onClick={() => setShowNewCategory(false)}>Bekor qilish</button>
              <button className="btn btn-primary" onClick={addCategory}><Plus size={14} /> Qo'shish</button>
            </div>
          </div>
        </div>
      )}

      {showSettings && (
        <div className="modal-backdrop" onClick={() => setShowSettings(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="card-title">Kassa sozlamalari</div>
            <div className="muted" style={{ fontSize: 12.5, marginBottom: 6 }}>Joriy kurs (1$ = ... so'm):</div>
            <div className="form-row wrap" style={{ marginBottom: 16 }}>
              <input className="input xs mono" type="number" value={usdRate} onChange={(e) => setUsdRate(e.target.value)} />
              <button className="btn btn-ghost" disabled={savingRate} onClick={saveRate}><Check size={14} /> Saqlash</button>
            </div>
            <div className="muted" style={{ fontSize: 12.5, marginBottom: 6 }}>Qaysi sanaga qoldiqni ko'rish ({currency === "usd" ? "Dollar" : "So'm"} kassa):</div>
            <div className="form-row wrap">
              <input className="input xs" type="date" value={asOfDate} onChange={(e) => setAsOfDate(e.target.value)} />
              {asOfDate && <span className="mono" style={{ fontWeight: 700 }}>{fmtNum(asOfBalance)} {symbol}</span>}
              {asOfDate && <button className="btn btn-ghost" onClick={() => setAsOfDate("")}>Tozalash</button>}
            </div>
            <div className="invoice-footer" style={{ marginTop: 16 }}>
              <button className="btn btn-primary" onClick={() => setShowSettings(false)}>Yopish</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- Products ---------------- */
function buildStockReport(products, movements, from, to) {
  return products.map((p) => {
    const pm = movements.filter((m) => m.product_id === p.id);
    const opening = from ? pm.filter((m) => m.date < from).reduce((s, m) => s + Number(m.qty), 0) : 0;
    const inRange = pm.filter((m) => (!from || m.date >= from) && (!to || m.date <= to));
    const kirim = inRange.filter((m) => Number(m.qty) > 0).reduce((s, m) => s + Number(m.qty), 0);
    const chiqim = inRange.filter((m) => Number(m.qty) < 0).reduce((s, m) => s + Math.abs(Number(m.qty)), 0);
    const closing = opening + kirim - chiqim;
    return { product: p, opening, kirim, chiqim, closing };
  });
}

function firstDayOfMonth() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-01`;
}

function Products({ products, movements, rawMaterials, rawMaterialBatches, productionBatches, productionConsumptions, suppliers, onReload }) {
  const [section, setSection] = useState("raw");
  const [rawBusy, setRawBusy] = useState(false);
  return (
    <div>
      <div className="tabs" style={{ padding: "0 0 10px", border: "none", marginBottom: 4 }}>
        <button className={`tab ${section === "raw" ? "active" : ""}`} onClick={() => setSection("raw")}>Xom ashyo ombori</button>
        <button className={`tab ${section === "finished" ? "active" : ""}`} onClick={() => setSection("finished")}>Tayyor mahsulot ombori</button>
      </div>
      {section === "raw" && (
        <RawMaterialsSection
          rawMaterials={rawMaterials} rawMaterialBatches={rawMaterialBatches}
          productionBatches={productionBatches} productionConsumptions={productionConsumptions}
          suppliers={suppliers}
          onReload={onReload} busy={rawBusy} setBusy={setRawBusy}
        />
      )}
      {section === "finished" && <FinishedGoodsWarehouse products={products} movements={movements} onReload={onReload} />}
    </div>
  );
}

function FinishedGoodsWarehouse({ products, movements, onReload }) {
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState(false);
  const [reportFrom, setReportFrom] = useState(firstDayOfMonth());
  const [reportTo, setReportTo] = useState(todayISO());
  const [showReport, setShowReport] = useState(false);

  const startEdit = (p) => { setEditingId(p.id); setEditForm({ ...p }); };
  const saveEdit = async () => {
    setBusy(true);
    try {
      await api.products.update(editingId, { name: editForm.name, article: editForm.article, unit: editForm.unit, price: Number(editForm.price) || 0 });
      setEditingId(null);
      await onReload();
    } finally { setBusy(false); }
  };
  const remove = async (id) => {
    if (!confirm("Bu mahsulotni o'chirasizmi? (Ishlab chiqarish bo'limidagi bog'lanish ham uzilib qoladi)")) return;
    setBusy(true);
    try { await api.products.remove(id); await onReload(); } finally { setBusy(false); }
  };

  const filtered = products.filter((p) => (p.name + (p.article || "")).toLowerCase().includes(query.toLowerCase()));
  const report = buildStockReport(products, movements, reportFrom, reportTo);

  return (
    <div>
      <div className="card">
        <div className="muted" style={{ fontSize: 12.5, marginBottom: 4 }}>
          Bu ombor faqat "Ishlab chiqarish" bo'limidan avtomatik to'ldiriladi. Yangi mahsulot turi qo'shish uchun "Ishlab chiqarish → Tayyor mahsulot" bo'limiga o'ting.
        </div>
      </div>

      <div className="card">
        <div className="card-title-row">
          <div className="card-title">Ombordagi mahsulotlar ({products.length})</div>
          <div className="search-box"><Search size={14} /><input placeholder="Qidirish…" value={query} onChange={(e) => setQuery(e.target.value)} /></div>
        </div>
        <table className="table">
          <thead><tr><th>Nomi</th><th>Artikul</th><th>Birlik</th><th className="right">Sotish narxi</th><th className="right">Tannarx</th><th className="right">Qoldiq</th><th className="right">Qiymati</th><th></th></tr></thead>
          <tbody>
            {filtered.map((p) => editingId === p.id ? (
              <tr key={p.id} className="editing-row">
                <td><input className="input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} /></td>
                <td><input className="input" value={editForm.article || ""} onChange={(e) => setEditForm({ ...editForm, article: e.target.value })} /></td>
                <td><input className="input" value={editForm.unit} onChange={(e) => setEditForm({ ...editForm, unit: e.target.value })} /></td>
                <td><input className="input mono" type="number" value={editForm.price} onChange={(e) => setEditForm({ ...editForm, price: e.target.value })} /></td>
                <td className="right mono muted">{fmt(p.cost_price || 0)}</td>
                <td className="right mono">{editForm.qty}</td>
                <td className="right mono">—</td>
                <td className="row-actions"><button className="icon-btn" disabled={busy} onClick={saveEdit}><Check size={15} /></button></td>
              </tr>
            ) : (
              <tr key={p.id} className={Number(p.qty) <= 5 ? "row-warn" : ""}>
                <td>{p.name}</td>
                <td className="muted">{p.article || "—"}</td>
                <td className="muted">{p.unit}</td>
                <td className="right mono">{fmt(p.price)}</td>
                <td className="right mono muted">{fmt(p.cost_price || 0)}</td>
                <td className="right mono">{p.qty}</td>
                <td className="right mono">{fmt(Number(p.qty) * Number(p.cost_price > 0 ? p.cost_price : p.price))}</td>
                <td className="row-actions">
                  <button className="icon-btn" onClick={() => startEdit(p)}><Pencil size={14} /></button>
                  <button className="icon-btn danger" disabled={busy} onClick={() => remove(p.id)}><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={8} className="empty">Mahsulot topilmadi</td></tr>}
          </tbody>
        </table>
      </div>

      <div className="card">
        <div className="card-title-row">
          <div className="card-title">Davr bo'yicha ombor harakati</div>
          <button className="btn btn-ghost" onClick={() => setShowReport(!showReport)}>{showReport ? "Yashirish" : "Ko'rsatish"}</button>
        </div>
        {showReport && (
          <>
            <div className="form-row wrap" style={{ marginBottom: 12 }}>
              <label className="muted" style={{ fontSize: 12.5 }}>Dan:</label>
              <input className="input xs" type="date" value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} />
              <label className="muted" style={{ fontSize: 12.5 }}>Gacha:</label>
              <input className="input xs" type="date" value={reportTo} onChange={(e) => setReportTo(e.target.value)} />
            </div>
            <table className="table">
              <thead><tr><th>Mahsulot</th><th className="right">Boshiga qoldiq</th><th className="right">Kirim</th><th className="right">Chiqim</th><th className="right">Oxiriga qoldiq</th></tr></thead>
              <tbody>
                {report.map((r) => (
                  <tr key={r.product.id}>
                    <td>{r.product.name}</td>
                    <td className="right mono">{r.opening} {r.product.unit}</td>
                    <td className="right mono tone-ok-text">{r.kirim ? `+${r.kirim}` : "0"} {r.product.unit}</td>
                    <td className="right mono tone-debt-text">{r.chiqim ? `-${r.chiqim}` : "0"} {r.product.unit}</td>
                    <td className="right mono"><b>{r.closing} {r.product.unit}</b></td>
                  </tr>
                ))}
                {report.length === 0 && <tr><td colSpan={5} className="empty">Mahsulot yo'q</td></tr>}
              </tbody>
            </table>
          </>
        )}
      </div>
    </div>
  );
}

/* ---------------- Ishlab chiqarish (Production) ---------------- */
function planFifoConsumption(rawMaterialId, requiredQty, batches) {
  const list = batches
    .filter((b) => b.raw_material_id === rawMaterialId && b.active && Number(b.remaining_qty) > 0.0000001)
    .sort((a, b) => new Date(a.date) - new Date(b.date) || new Date(a.created_at) - new Date(b.created_at));
  const plan = [];
  let remaining = requiredQty;
  for (const b of list) {
    if (remaining <= 0.0000001) break;
    const take = Math.min(remaining, Number(b.remaining_qty));
    plan.push({ batch: b, qty: take });
    remaining -= take;
  }
  return { plan, shortfall: Math.max(remaining, 0) };
}

function computeProductionPlan(finishedProductId, qty, norms, batches, rawMaterials) {
  const relevantNorms = norms.filter((n) => n.finished_product_id === finishedProductId);
  const details = [];
  let totalCost = 0;
  const errors = [];
  for (const norm of relevantNorms) {
    const required = Number(norm.qty_per_unit) * Number(qty);
    const { plan, shortfall } = planFifoConsumption(norm.raw_material_id, required, batches);
    const material = rawMaterials.find((m) => m.id === norm.raw_material_id);
    const cost = plan.reduce((s, p) => s + p.qty * Number(p.batch.unit_cost), 0);
    totalCost += cost;
    details.push({ material, required, plan, cost });
    if (shortfall > 0.0000001) {
      errors.push(`"${material?.name || "?"}" yetishmayapti: kerak ${required.toFixed(1)}, yetmadi ${shortfall.toFixed(1)} ${material?.unit || ""}`);
    }
  }
  return { details, totalCost, unitCost: qty > 0 ? totalCost / qty : 0, errors, hasNorms: relevantNorms.length > 0 };
}

const OVERHEAD_EXCLUDE_CATEGORIES = ["Xom ashyo xaridi", "Sotuvdan tushum", "Kapital kiritish", "Boshqa tushum"];
function computeOverheadPerLiter(cashTransactions, productionBatches, finishedProducts, fixedAssets = [], days = 30) {
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const cutoffStr = cutoff.toISOString().slice(0, 10);
  const cashOpex = cashTransactions
    .filter((t) => t.type === "chiqim" && t.date >= cutoffStr && !OVERHEAD_EXCLUDE_CATEGORIES.includes(t.category))
    .reduce((s, t) => s + Number(t.amount), 0);
  const monthlyDepreciation = fixedAssets
    .filter((a) => a.include_in_cost)
    .reduce((s, a) => s + (Number(a.purchase_price) || 0) * (Number(a.depreciation_rate) || 0) / 100 / 12, 0);
  const depreciationForPeriod = monthlyDepreciation * (days / 30);
  const opex = cashOpex + depreciationForPeriod;
  const totalLiters = productionBatches
    .filter((b) => b.date >= cutoffStr)
    .reduce((s, b) => {
      const fp = finishedProducts.find((f) => f.id === b.finished_product_id);
      const vol = Number(fp?.volume_liters) || 1;
      return s + Number(b.qty) * vol;
    }, 0);
  const overheadPerLiter = totalLiters > 0 ? opex / totalLiters : 0;
  return { overheadPerLiter, opex, cashOpex, depreciationForPeriod, totalLiters, days };
}

function Production({ products, finishedProducts, rawMaterials, rawMaterialBatches, productNorms, productionBatches, productionConsumptions, cashTransactions, currentUser, onReload, prefill, onPrefillConsumed, fixedAssets }) {
  const [section, setSection] = useState(prefill ? "produce" : "produce");
  const [busy, setBusy] = useState(false);

  return (
    <div>
      <div className="tabs" style={{ padding: "0 0 10px", border: "none", marginBottom: 4 }}>
        {[
          { id: "produce", label: "Ishlab chiqarish" },
          { id: "finished", label: "Tayyor mahsulot" },
          { id: "norms", label: "Norma rasxod" },
        ].map((s) => (
          <button key={s.id} className={`tab ${section === s.id ? "active" : ""}`} onClick={() => setSection(s.id)}>{s.label}</button>
        ))}
      </div>
      {section === "finished" && <FinishedProductsSection finishedProducts={finishedProducts} products={products} onReload={onReload} busy={busy} setBusy={setBusy} />}
      {section === "norms" && <NormsSection finishedProducts={finishedProducts} rawMaterials={rawMaterials} productNorms={productNorms} onReload={onReload} busy={busy} setBusy={setBusy} />}
      {section === "produce" && (
        <ProduceSection
          finishedProducts={finishedProducts} products={products} rawMaterials={rawMaterials}
          rawMaterialBatches={rawMaterialBatches} productNorms={productNorms} productionBatches={productionBatches}
          productionConsumptions={productionConsumptions} cashTransactions={cashTransactions} currentUser={currentUser}
          onReload={onReload} busy={busy} setBusy={setBusy}
          prefill={prefill} onPrefillConsumed={onPrefillConsumed} fixedAssets={fixedAssets}
        />
      )}
    </div>
  );
}

function buildRawMaterialReport(rawMaterials, rawMaterialBatches, productionConsumptions, productionBatches, from, to) {
  const batchToMaterial = {};
  rawMaterialBatches.forEach((b) => { batchToMaterial[b.id] = b.raw_material_id; });
  const batchDateMap = {};
  productionBatches.forEach((pb) => { batchDateMap[pb.id] = pb.date; });

  return rawMaterials.map((m) => {
    const kirimEvents = rawMaterialBatches
      .filter((b) => b.raw_material_id === m.id && !b.is_adjustment)
      .map((b) => ({ date: b.date, qty: Number(b.qty) }));
    const chiqimEvents = productionConsumptions
      .filter((c) => batchToMaterial[c.raw_material_batch_id] === m.id)
      .map((c) => ({ date: batchDateMap[c.production_batch_id], qty: Number(c.qty) }))
      .filter((e) => e.date);

    const opening = from
      ? kirimEvents.filter((e) => e.date < from).reduce((s, e) => s + e.qty, 0) - chiqimEvents.filter((e) => e.date < from).reduce((s, e) => s + e.qty, 0)
      : 0;
    const kirim = kirimEvents.filter((e) => (!from || e.date >= from) && (!to || e.date <= to)).reduce((s, e) => s + e.qty, 0);
    const chiqim = chiqimEvents.filter((e) => (!from || e.date >= from) && (!to || e.date <= to)).reduce((s, e) => s + e.qty, 0);
    const closing = opening + kirim - chiqim;
    return { material: m, opening, kirim, chiqim, closing };
  });
}

function RawMaterialsSection({ rawMaterials, rawMaterialBatches, productionBatches, productionConsumptions, suppliers, onReload, busy, setBusy }) {
  const [showAddMaterial, setShowAddMaterial] = useState(false);
  const [form, setForm] = useState({ name: "", unit: "kg" });
  const [batchForm, setBatchForm] = useState({ raw_material_id: "", qty: "", unit_cost: "", date: todayISO(), note: "", supplier_id: "" });
  const [freezeFor, setFreezeFor] = useState(null);
  const [freezeQty, setFreezeQty] = useState("");
  const [reportFrom, setReportFrom] = useState(firstDayOfMonth());
  const [reportTo, setReportTo] = useState(todayISO());
  const [showReport, setShowReport] = useState(false);
  const [batchError, setBatchError] = useState("");
  const [expandedMaterials, setExpandedMaterials] = useState({});
  const toggleMaterial = (id) => setExpandedMaterials((prev) => ({ ...prev, [id]: !prev[id] }));

  const addMaterial = async () => {
    if (!form.name.trim()) return;
    setBusy(true);
    try { await api.rawMaterials.add({ name: form.name.trim(), unit: form.unit || "kg" }); setForm({ name: "", unit: "kg" }); setShowAddMaterial(false); await onReload(); } finally { setBusy(false); }
  };
  const removeMaterial = async (id) => {
    if (!confirm("Bu xom ashyo turini o'chirasizmi? (Barcha partiyalari ham o'chadi)")) return;
    setBusy(true);
    try { await api.rawMaterials.remove(id); await onReload(); } finally { setBusy(false); }
  };

  const addBatch = async () => {
    setBatchError("");
    if (!batchForm.raw_material_id || !Number(batchForm.qty)) return;
    if (!batchForm.unit_cost || Number(batchForm.unit_cost) <= 0) { setBatchError("Xom ashyo narxini kiritish majburiy"); return; }
    if (!batchForm.supplier_id) { setBatchError("Ta'minotchini tanlash majburiy"); return; }
    setBusy(true);
    try {
      const qty = Number(batchForm.qty);
      await api.rawMaterialBatches.add({
        raw_material_id: batchForm.raw_material_id, date: batchForm.date, qty, remaining_qty: qty,
        unit_cost: Number(batchForm.unit_cost), active: true, note: batchForm.note, supplier_id: batchForm.supplier_id,
      });
      setBatchForm({ raw_material_id: "", qty: "", unit_cost: "", date: todayISO(), note: "", supplier_id: "" });
      await onReload();
    } finally { setBusy(false); }
  };

  const openFreeze = (b) => { setFreezeFor(b); setFreezeQty(String(b.remaining_qty)); };
  const submitFreeze = async () => {
    const q = Number(freezeQty);
    if (!q || q <= 0 || q > Number(freezeFor.remaining_qty)) return;
    setBusy(true);
    try {
      await api.rawMaterialBatches.update(freezeFor.id, { remaining_qty: Number(freezeFor.remaining_qty) - q });
      await api.rawMaterialBatches.add({
        raw_material_id: freezeFor.raw_material_id, date: freezeFor.date, qty: q, remaining_qty: q,
        unit_cost: freezeFor.unit_cost, active: false, note: "Harakatsiz qilindi", is_adjustment: true,
      });
      setFreezeFor(null);
      await onReload();
    } finally { setBusy(false); }
  };

  const reactivate = async (b) => {
    setBusy(true);
    try { await api.rawMaterialBatches.update(b.id, { active: true, date: todayISO(), is_adjustment: false }); await onReload(); } finally { setBusy(false); }
  };

  const report = buildRawMaterialReport(rawMaterials, rawMaterialBatches, productionConsumptions, productionBatches, reportFrom, reportTo);
  const totalsAsOfToday = buildRawMaterialReport(rawMaterials, rawMaterialBatches, productionConsumptions, productionBatches, "", todayISO());

  const removeBatch = async (b) => {
    if (!confirm("Bu partiyani butunlay o'chirasizmi? Bu amalni qaytarib bo'lmaydi.")) return;
    setBusy(true);
    try { await api.rawMaterialBatches.remove(b.id); await onReload(); } finally { setBusy(false); }
  };

  return (
    <div>
      <div className="card">
        <div className="card-title-row">
          <div className="card-title">Xom ashyo kirim qilish (FIFO partiya)</div>
          <button className="btn btn-primary" onClick={() => setShowAddMaterial(true)}><Plus size={14} /> Yangi xom ashyo turi</button>
        </div>
        <div className="form-row wrap">
          <select className="input" value={batchForm.raw_material_id} onChange={(e) => setBatchForm({ ...batchForm, raw_material_id: e.target.value })}>
            <option value="">— Xom ashyo tanlang —</option>
            {rawMaterials.map((m) => <option key={m.id} value={m.id}>{m.name} ({m.unit})</option>)}
          </select>
          <input className="input sm mono" type="number" placeholder="Miqdor" value={batchForm.qty} onChange={(e) => setBatchForm({ ...batchForm, qty: e.target.value })} />
          <input className="input sm mono" type="number" placeholder="Narxi (majburiy)" value={batchForm.unit_cost} onChange={(e) => setBatchForm({ ...batchForm, unit_cost: e.target.value })} />
          <input className="input xs" type="date" value={batchForm.date} onChange={(e) => setBatchForm({ ...batchForm, date: e.target.value })} />
          <select className="input" value={batchForm.supplier_id} onChange={(e) => setBatchForm({ ...batchForm, supplier_id: e.target.value })}>
            <option value="">— Ta'minotchini tanlang (majburiy) —</option>
            {suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <button className="btn btn-primary" disabled={busy} onClick={addBatch}><Plus size={14} /> Kirim qilish</button>
        </div>
        {batchError && <div className="error-box" style={{ marginTop: 8 }}><AlertTriangle size={14} /> {batchError}</div>}
      </div>

      <div className="card">
        <div className="card-title-row">
          <div className="card-title">Xom ashyo partiyalari (FIFO)</div>
          <div className="form-row wrap">
            <button className="btn btn-ghost" onClick={() => setExpandedMaterials(Object.fromEntries(rawMaterials.map((m) => [m.id, true])))}>Barchasini ochish</button>
            <button className="btn btn-ghost" onClick={() => setExpandedMaterials({})}>Barchasini yig'ish</button>
          </div>
        </div>
        {rawMaterials.map((m) => {
          const batches = rawMaterialBatches.filter((b) => b.raw_material_id === m.id).sort((a, b) => new Date(a.date) - new Date(b.date));
          const totalRemaining = totalsAsOfToday.find((r) => r.material.id === m.id)?.closing ?? 0;
          const activeRemaining = batches.filter((b) => b.active).reduce((s, b) => s + Number(b.remaining_qty), 0);
          if (batches.length === 0) return null;
          const isOpen = !!expandedMaterials[m.id];
          return (
            <div key={m.id} style={{ marginBottom: 10 }}>
              <div className="card-title-row clickable-row" style={{ marginBottom: isOpen ? 6 : 0, padding: "6px 4px" }} onClick={() => toggleMaterial(m.id)}>
                <b>{isOpen ? "▾" : "▸"} {m.name} <span className="muted" style={{ fontWeight: 400 }}>({batches.length} partiya)</span></b>
                <span className="muted">Umumiy qoldiq: <span className={`mono ${totalRemaining < 0 ? "tone-debt-text" : ""}`}>{fmtQty(totalRemaining)} {m.unit}</span>
                  {activeRemaining !== totalRemaining && <span style={{ fontSize: 11 }}> (shundan faol: {fmtQty(activeRemaining)} {m.unit})</span>}
                </span>
              </div>
              {isOpen && (
                <table className="table">
                  <thead><tr><th>Sana</th><th className="right">Kelgan</th><th className="right">Qolgan</th><th className="right">Tannarx</th><th>Holat</th><th></th></tr></thead>
                  <tbody>
                    {batches.map((b) => (
                      <tr key={b.id} className={!b.active ? "muted" : Number(b.remaining_qty) < 0 ? "row-warn" : ""}>
                        <td>{fmtDate(b.date)}</td>
                        <td className="right mono">{fmtQty(b.qty)} {m.unit}</td>
                        <td className={`right mono ${Number(b.remaining_qty) < 0 ? "tone-debt-text" : ""}`}>
                          {fmtQty(b.remaining_qty)} {m.unit}
                          {Number(b.remaining_qty) < 0 && <div className="muted" style={{ fontSize: 10.5 }}>xato: manfiy qoldiq</div>}
                        </td>
                        <td className="right mono">{fmt(b.unit_cost)}</td>
                        <td>{b.active ? <span className="tone-ok-text">Faol</span> : <span className="tone-debt-text">Harakatsiz</span>}</td>
                        <td className="row-actions">
                          {b.active ? (
                            Number(b.remaining_qty) > 0 && <button className="btn btn-ghost" style={{ padding: "4px 8px" }} onClick={() => openFreeze(b)}><Lock size={13} /> Harakatsiz</button>
                          ) : (
                            <button className="btn btn-ghost" style={{ padding: "4px 8px" }} onClick={() => reactivate(b)}><Unlock size={13} /> Faollashtirish</button>
                          )}
                          {Number(b.remaining_qty) < 0 && (
                            <button className="btn btn-ghost" style={{ padding: "4px 8px" }} onClick={async () => { setBusy(true); try { await api.rawMaterialBatches.update(b.id, { remaining_qty: 0 }); await onReload(); } finally { setBusy(false); } }}>
                              <Check size={13} /> Nolga tuzatish
                            </button>
                          )}
                          <button className="icon-btn danger" onClick={() => removeBatch(b)}><Trash2 size={14} /></button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          );
        })}
        {rawMaterials.length === 0 && <div className="empty">Hali xom ashyo turi qo'shilmagan</div>}
      </div>

      <div className="card">
        <div className="card-title-row">
          <div className="card-title">Davr bo'yicha xom ashyo harakati</div>
          <button className="btn btn-ghost" onClick={() => setShowReport(!showReport)}>{showReport ? "Yashirish" : "Ko'rsatish"}</button>
        </div>
        {showReport && (
          <>
            <div className="form-row wrap" style={{ marginBottom: 12 }}>
              <label className="muted" style={{ fontSize: 12.5 }}>Dan:</label>
              <input className="input xs" type="date" value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} />
              <label className="muted" style={{ fontSize: 12.5 }}>Gacha:</label>
              <input className="input xs" type="date" value={reportTo} onChange={(e) => setReportTo(e.target.value)} />
            </div>
            <table className="table">
              <thead><tr><th>Xom ashyo</th><th className="right">Boshiga qoldiq</th><th className="right">Kirim</th><th className="right">Chiqim</th><th className="right">Oxiriga qoldiq</th></tr></thead>
              <tbody>
                {report.map((r) => (
                  <tr key={r.material.id}>
                    <td>{r.material.name}</td>
                    <td className="right mono">{fmtQty(r.opening)} {r.material.unit}</td>
                    <td className="right mono tone-ok-text">{r.kirim ? `+${fmtQty(r.kirim)}` : "0"} {r.material.unit}</td>
                    <td className="right mono tone-debt-text">{r.chiqim ? `-${fmtQty(r.chiqim)}` : "0"} {r.material.unit}</td>
                    <td className="right mono"><b>{fmtQty(r.closing)} {r.material.unit}</b></td>
                  </tr>
                ))}
                {report.length === 0 && <tr><td colSpan={5} className="empty">Xom ashyo yo'q</td></tr>}
              </tbody>
            </table>
          </>
        )}
      </div>

      {freezeFor && (
        <div className="modal-backdrop" onClick={() => setFreezeFor(null)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="card-title">Harakatsiz qilish</div>
            <div className="muted" style={{ marginBottom: 10, fontSize: 12.5 }}>
              Qolgan: {freezeFor.remaining_qty} — qancha qismini harakatsiz qilasiz?
            </div>
            <input className="input mono" type="number" value={freezeQty} onChange={(e) => setFreezeQty(e.target.value)} />
            <div className="invoice-footer" style={{ marginTop: 14 }}>
              <button className="btn btn-ghost" onClick={() => setFreezeFor(null)}>Bekor qilish</button>
              <button className="btn btn-primary" disabled={busy} onClick={submitFreeze}><Lock size={14} /> Tasdiqlash</button>
            </div>
          </div>
        </div>
      )}

      {showAddMaterial && (
        <div className="modal-backdrop" onClick={() => setShowAddMaterial(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="card-title">Yangi xom ashyo turi</div>
            <div className="form-row wrap" style={{ marginBottom: 8 }}>
              <input className="input" placeholder="Nomi (masalan: Paxta yog'i)" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </div>
            <div className="form-row wrap">
              <input className="input xs" placeholder="Birlik (kg, dona...)" value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} />
            </div>
            <div className="invoice-footer" style={{ marginTop: 14 }}>
              <button className="btn btn-ghost" onClick={() => setShowAddMaterial(false)}>Bekor qilish</button>
              <button className="btn btn-primary" disabled={busy} onClick={addMaterial}><Plus size={14} /> Qo'shish</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function FinishedProductsSection({ finishedProducts, products, onReload, busy, setBusy }) {
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState({ name: "", unit: "dona", linked_product_id: "", volume_liters: "1" });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});

  const add = async () => {
    if (!form.name.trim()) return;
    setBusy(true);
    try {
      let linkedId = form.linked_product_id || null;
      if (!linkedId) {
        const created = await api.products.add({ name: form.name.trim(), article: "", unit: form.unit || "dona", price: 0, qty: 0 });
        linkedId = created.id;
      }
      await api.finishedProducts.add({ name: form.name.trim(), unit: form.unit || "dona", linked_product_id: linkedId, volume_liters: Number(form.volume_liters) || 1 });
      setForm({ name: "", unit: "dona", linked_product_id: "", volume_liters: "1" });
      setShowAdd(false);
      await onReload();
    } finally { setBusy(false); }
  };
  const remove = async (id) => {
    if (!confirm("Bu tayyor mahsulotni o'chirasizmi?")) return;
    setBusy(true);
    try { await api.finishedProducts.remove(id); await onReload(); } finally { setBusy(false); }
  };
  const startEdit = (f) => { setEditingId(f.id); setEditForm({ ...f }); };
  const saveEdit = async () => {
    setBusy(true);
    try {
      await api.finishedProducts.update(editingId, { name: editForm.name, unit: editForm.unit, volume_liters: Number(editForm.volume_liters) || 1 });
      setEditingId(null);
      await onReload();
    } finally { setBusy(false); }
  };

  return (
    <div>
      <div className="card">
        <div className="card-title-row">
          <div className="card-title">Tayyor mahsulotlar ({finishedProducts.length})</div>
          <button className="btn btn-primary" onClick={() => setShowAdd(true)}><Plus size={14} /> Yangi tayyor mahsulot</button>
        </div>
        <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>
          "Litr hajmi" — operatsion xarajatlarni (ijara, ish haqi va h.k.) mahsulotlar orasida hajmga mutanosib taqsimlash uchun ishlatiladi (masalan 5L = 5, 1L = 1).
        </div>
        <table className="table">
          <thead><tr><th>Nomi</th><th>Birlik</th><th className="right">Litr hajmi</th><th>Bog'langan Ombor mahsuloti</th><th></th></tr></thead>
          <tbody>
            {finishedProducts.map((f) => editingId === f.id ? (
              <tr key={f.id} className="editing-row">
                <td><input className="input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} /></td>
                <td><input className="input" value={editForm.unit} onChange={(e) => setEditForm({ ...editForm, unit: e.target.value })} /></td>
                <td><input className="input mono" type="number" value={editForm.volume_liters} onChange={(e) => setEditForm({ ...editForm, volume_liters: e.target.value })} /></td>
                <td className="muted">{products.find((p) => p.id === f.linked_product_id)?.name || "— bog'lanmagan —"}</td>
                <td className="row-actions"><button className="icon-btn" disabled={busy} onClick={saveEdit}><Check size={15} /></button></td>
              </tr>
            ) : (
              <tr key={f.id}>
                <td>{f.name}</td>
                <td className="muted">{f.unit}</td>
                <td className="right mono">{fmtQty(f.volume_liters || 1)} L</td>
                <td className="muted">{products.find((p) => p.id === f.linked_product_id)?.name || "— bog'lanmagan —"}</td>
                <td className="row-actions">
                  <button className="icon-btn" onClick={() => startEdit(f)}><Pencil size={14} /></button>
                  <button className="icon-btn danger" onClick={() => remove(f.id)}><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
            {finishedProducts.length === 0 && <tr><td colSpan={5} className="empty">Hali tayyor mahsulot qo'shilmagan</td></tr>}
          </tbody>
        </table>
      </div>

      {showAdd && (
        <div className="modal-backdrop" onClick={() => setShowAdd(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="card-title">Yangi tayyor mahsulot</div>
            <div className="muted" style={{ marginBottom: 10, fontSize: 12.5 }}>
              Bo'sh qoldirsangiz, Ombor'ga shu nom bilan avtomatik yangi mahsulot yaratiladi. Mavjud mahsulotga bog'lamoqchi bo'lsangiz, pastdan tanlang.
            </div>
            <div className="form-row wrap" style={{ marginBottom: 8 }}>
              <input className="input" placeholder="Nomi (masalan: 5L yog')" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
              <input className="input xs" placeholder="Birlik" value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} />
              <input className="input xs mono" type="number" placeholder="Litr hajmi" value={form.volume_liters} onChange={(e) => setForm({ ...form, volume_liters: e.target.value })} />
            </div>
            <div className="form-row wrap">
              <select className="input" value={form.linked_product_id} onChange={(e) => setForm({ ...form, linked_product_id: e.target.value })}>
                <option value="">— Yangi Ombor mahsuloti yaratiladi —</option>
                {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </div>
            <div className="invoice-footer" style={{ marginTop: 14 }}>
              <button className="btn btn-ghost" onClick={() => setShowAdd(false)}>Bekor qilish</button>
              <button className="btn btn-primary" disabled={busy} onClick={add}><Plus size={14} /> Qo'shish</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function NormsSection({ finishedProducts, rawMaterials, productNorms, onReload, busy, setBusy }) {
  const [selected, setSelected] = useState("");
  const [form, setForm] = useState({ raw_material_id: "", qty_per_unit: "" });

  const addNorm = async () => {
    if (!selected || !form.raw_material_id || !Number(form.qty_per_unit)) return;
    setBusy(true);
    try {
      await api.productNorms.add({ finished_product_id: selected, raw_material_id: form.raw_material_id, qty_per_unit: Number(form.qty_per_unit) });
      setForm({ raw_material_id: "", qty_per_unit: "" });
      await onReload();
    } finally { setBusy(false); }
  };
  const removeNorm = async (id) => {
    setBusy(true);
    try { await api.productNorms.remove(id); await onReload(); } finally { setBusy(false); }
  };

  const norms = productNorms.filter((n) => n.finished_product_id === selected);

  return (
    <div>
      <div className="card">
        <div className="card-title">Norma rasxod jadvali</div>
        <div style={{ overflowX: "auto" }}>
          <table className="table">
            <thead>
              <tr>
                <th>Tovar artikuli</th>
                {rawMaterials.map((m) => <th key={m.id} className="right">{m.name}</th>)}
              </tr>
            </thead>
            <tbody>
              {finishedProducts.map((f) => (
                <tr key={f.id}>
                  <td>{f.name} <span className="muted">({f.unit})</span></td>
                  {rawMaterials.map((m) => {
                    const norm = productNorms.find((n) => n.finished_product_id === f.id && n.raw_material_id === m.id);
                    return <td key={m.id} className="right mono">{norm ? norm.qty_per_unit : "—"}</td>;
                  })}
                </tr>
              ))}
              {finishedProducts.length === 0 && <tr><td colSpan={rawMaterials.length + 1} className="empty">Hali tayyor mahsulot yo'q</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      <div className="card">
      <div className="card-title">Norma rasxodni tahrirlash</div>
      <div className="form-row wrap" style={{ marginBottom: 14 }}>
        <select className="input" value={selected} onChange={(e) => setSelected(e.target.value)}>
          <option value="">— Tayyor mahsulotni tanlang —</option>
          {finishedProducts.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
        </select>
      </div>

      {selected && (
        <>
          <table className="table" style={{ marginBottom: 12 }}>
            <thead><tr><th>Xom ashyo</th><th className="right">1 dona uchun</th><th></th></tr></thead>
            <tbody>
              {norms.map((n) => {
                const m = rawMaterials.find((x) => x.id === n.raw_material_id);
                return (
                  <tr key={n.id}>
                    <td>{m?.name || "?"}</td>
                    <td className="right mono">{n.qty_per_unit} {m?.unit}</td>
                    <td className="row-actions"><button className="icon-btn danger" onClick={() => removeNorm(n.id)}><Trash2 size={14} /></button></td>
                  </tr>
                );
              })}
              {norms.length === 0 && <tr><td colSpan={3} className="empty">Norma belgilanmagan</td></tr>}
            </tbody>
          </table>
          <div className="form-row wrap">
            <select className="input" value={form.raw_material_id} onChange={(e) => setForm({ ...form, raw_material_id: e.target.value })}>
              <option value="">— Xom ashyo tanlang —</option>
              {rawMaterials.map((m) => <option key={m.id} value={m.id}>{m.name} ({m.unit})</option>)}
            </select>
            <input className="input sm mono" type="number" placeholder="Miqdor" value={form.qty_per_unit} onChange={(e) => setForm({ ...form, qty_per_unit: e.target.value })} />
            <button className="btn btn-primary" disabled={busy} onClick={addNorm}><Plus size={14} /> Qo'shish</button>
          </div>
        </>
      )}
      </div>
    </div>
  );
}

function computeMultiProductionPlan(rows, productNorms, rawMaterialBatches, rawMaterials, finishedProducts) {
  let working = rawMaterialBatches.map((b) => ({ ...b }));
  const rowResults = [];
  for (const row of rows) {
    if (!row.finishedId || !Number(row.qty)) continue;
    const fp = finishedProducts.find((f) => f.id === row.finishedId);
    const plan = computeProductionPlan(row.finishedId, Number(row.qty), productNorms, working, rawMaterials);
    rowResults.push({ row, finished: fp, plan });
    if (plan.hasNorms && plan.errors.length === 0) {
      for (const d of plan.details) {
        for (const p of d.plan) {
          const wb = working.find((b) => b.id === p.batch.id);
          if (wb) wb.remaining_qty = Number(wb.remaining_qty) - p.qty;
        }
      }
    }
  }
  return rowResults;
}

function ProduceSection({ finishedProducts, products, rawMaterials, rawMaterialBatches, productNorms, productionBatches, productionConsumptions, cashTransactions, currentUser, onReload, busy, setBusy, prefill, onPrefillConsumed, fixedAssets }) {
  const [rows, setRows] = useState([{ id: uid(), finishedId: "", qty: "" }]);
  const [date, setDate] = useState(todayISO());
  const [preview, setPreview] = useState(null); // array of { row, finished, plan }
  const [error, setError] = useState("");

  useEffect(() => {
    if (!prefill) return;
    const fp = finishedProducts.find((f) => f.linked_product_id === prefill.productId);
    if (fp) {
      setRows([{ id: uid(), finishedId: fp.id, qty: String(Math.ceil(Number(prefill.qty))) }]);
      setError(""); setPreview(null);
    } else {
      const p = products.find((x) => x.id === prefill.productId);
      setError(`"${p?.name || "Mahsulot"}" uchun tayyor mahsulot/norma rasxod bog'lanmagan — avval "Tayyor mahsulot" bo'limida bog'lang.`);
    }
    if (onPrefillConsumed) onPrefillConsumed();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [prefill]);

  const overhead = computeOverheadPerLiter(cashTransactions, productionBatches, finishedProducts, fixedAssets, 30);
  const overheadFor = (finished) => overhead.overheadPerLiter * (Number(finished?.volume_liters) || 1);

  const addRow = () => setRows([...rows, { id: uid(), finishedId: "", qty: "" }]);
  const updateRow = (id, patch) => { setRows(rows.map((r) => (r.id === id ? { ...r, ...patch } : r))); setPreview(null); };
  const removeRow = (id) => setRows(rows.filter((r) => r.id !== id));

  const calculate = () => {
    setError(""); setPreview(null);
    const activeRows = rows.filter((r) => r.finishedId && Number(r.qty));
    if (activeRows.length === 0) { setError("Kamida bitta mahsulot va miqdorni kiriting"); return; }
    const results = computeMultiProductionPlan(activeRows, productNorms, rawMaterialBatches, rawMaterials, finishedProducts);
    const noNorms = results.filter((r) => !r.plan.hasNorms);
    if (noNorms.length > 0) {
      setError(noNorms.map((r) => `"${r.finished?.name}" uchun norma rasxod belgilanmagan`).join("; "));
      return;
    }
    setPreview(results);
  };

  const confirmProduce = async () => {
    if (!preview) return;
    const activeRows = rows.filter((r) => r.finishedId && Number(r.qty));
    // Tasdiqlashdan oldin joriy ombor holati bo'yicha qayta tekshiramiz (eski/eskirgan hisob-kitobning oldini olish uchun)
    const fresh = computeMultiProductionPlan(activeRows, productNorms, rawMaterialBatches, rawMaterials, finishedProducts);
    const errs = fresh.filter((r) => r.plan.errors.length > 0);
    if (errs.length > 0) {
      setError("Ombor holati o'zgargan, qayta hisoblang: " + errs.flatMap((r) => r.plan.errors).join("; "));
      setPreview(null);
      return;
    }
    setBusy(true);
    try {
      for (const { row, finished, plan } of fresh) {
        const rowOverhead = overheadFor(finished);
        const finalUnitCost = plan.unitCost + rowOverhead;
        const finalTotalCost = finalUnitCost * Number(row.qty);
        const batch = await api.productionBatches.add({
          finished_product_id: row.finishedId, date, qty: Number(row.qty), unit_cost: finalUnitCost, total_cost: finalTotalCost, created_by: currentUser || null,
        });
        for (const d of plan.details) {
          for (const p of d.plan) {
            await api.productionConsumptions.add({ production_batch_id: batch.id, raw_material_batch_id: p.batch.id, qty: p.qty, unit_cost: p.batch.unit_cost });
            await api.rawMaterialBatches.update(p.batch.id, { remaining_qty: Math.max(0, Number(p.batch.remaining_qty) - p.qty) });
          }
        }
        if (finished?.linked_product_id) {
          const linkedProduct = products.find((pp) => pp.id === finished.linked_product_id);
          if (linkedProduct) {
            await api.products.update(linkedProduct.id, { qty: Number(linkedProduct.qty) + Number(row.qty), cost_price: finalUnitCost });
            await api.stockMovements.add({ product_id: linkedProduct.id, date, qty: Number(row.qty), type: "ishlab_chiqarish", note: `Ishlab chiqarish (tannarx: ${fmt(finalUnitCost)}, shundan xarajat ulushi: ${fmt(rowOverhead)})` });
          }
        }
      }
      setPreview(null); setRows([{ id: uid(), finishedId: "", qty: "" }]);
      await onReload();
    } catch (e) {
      setError(e.message || "Xatolik yuz berdi");
    } finally {
      setBusy(false);
    }
  };

  const deleteProduction = async (batch) => {
    if (!confirm("Bu ishlab chiqarishni bekor qilasizmi? Sarflangan xom ashyo omborga qaytariladi va tayyor mahsulot qoldig'idan ayiriladi.")) return;
    setBusy(true);
    try {
      const consumptions = productionConsumptions.filter((c) => c.production_batch_id === batch.id);
      for (const c of consumptions) {
        const rmBatch = rawMaterialBatches.find((b) => b.id === c.raw_material_batch_id);
        if (rmBatch) {
          await api.rawMaterialBatches.update(rmBatch.id, { remaining_qty: Number(rmBatch.remaining_qty) + Number(c.qty) });
        }
      }
      await api.productionConsumptions.removeByBatch(batch.id);

      const finishedProd = finishedProducts.find((f) => f.id === batch.finished_product_id);
      if (finishedProd?.linked_product_id) {
        const linkedProduct = products.find((p) => p.id === finishedProd.linked_product_id);
        if (linkedProduct) {
          await api.products.update(linkedProduct.id, { qty: Number(linkedProduct.qty) - Number(batch.qty) });
          await api.stockMovements.add({ product_id: linkedProduct.id, date: todayISO(), qty: -Number(batch.qty), type: "tuzatish", note: "Ishlab chiqarish bekor qilindi" });
        }
      }

      await api.productionBatches.remove(batch.id);
      await onReload();
    } finally {
      setBusy(false);
    }
  };

  const recalcProduction = async (batch) => {
    if (!confirm("Bu ishlab chiqarishni JORIY norma rasxod va joriy xarajat ulushi asosida qayta hisoblaysizmi? Eski sarflangan xom ashyo omborga qaytariladi va joriy norma bo'yicha qaytadan sarflanadi. Ishlab chiqarilgan mahsulot miqdori o'zgarmaydi, faqat sarflangan xom ashyo va tannarx to'g'rilanadi.")) return;
    setBusy(true);
    try {
      const oldConsumptions = productionConsumptions.filter((c) => c.production_batch_id === batch.id);
      const finishedProd = finishedProducts.find((f) => f.id === batch.finished_product_id);

      // Eski sarflangan xom ashyoni (xotirada) qaytarilgan holatda hisoblab, shu asosda reja tuzamiz
      const restoredBatches = rawMaterialBatches.map((b) => {
        const consumed = oldConsumptions.filter((c) => c.raw_material_batch_id === b.id).reduce((s, c) => s + Number(c.qty), 0);
        return consumed ? { ...b, remaining_qty: Number(b.remaining_qty) + consumed } : b;
      });

      const plan = computeProductionPlan(batch.finished_product_id, Number(batch.qty), productNorms, restoredBatches, rawMaterials);
      if (!plan.hasNorms) { alert("Bu mahsulot uchun norma rasxod topilmadi."); return; }
      if (plan.errors.length > 0) { alert("Qayta hisoblab bo'lmadi — yetarli xom ashyo yo'q:\n" + plan.errors.join("\n")); return; }

      const finalUnitCost = plan.unitCost + overheadFor(finishedProd);
      const finalTotalCost = finalUnitCost * Number(batch.qty);

      // Bazada: avval eski sarfni qaytarish
      for (const c of oldConsumptions) {
        const rmBatch = rawMaterialBatches.find((b) => b.id === c.raw_material_batch_id);
        if (rmBatch) await api.rawMaterialBatches.update(rmBatch.id, { remaining_qty: Number(rmBatch.remaining_qty) + Number(c.qty) });
      }
      await api.productionConsumptions.removeByBatch(batch.id);

      // Yangi (to'g'irlangan) sarfni qo'llash
      for (const d of plan.details) {
        for (const p of d.plan) {
          await api.productionConsumptions.add({ production_batch_id: batch.id, raw_material_batch_id: p.batch.id, qty: p.qty, unit_cost: p.batch.unit_cost });
          await api.rawMaterialBatches.update(p.batch.id, { remaining_qty: Math.max(0, Number(p.batch.remaining_qty) - p.qty) });
        }
      }

      await api.productionBatches.update(batch.id, { unit_cost: finalUnitCost, total_cost: finalTotalCost });

      if (finishedProd?.linked_product_id) {
        const linkedProduct = products.find((p) => p.id === finishedProd.linked_product_id);
        if (linkedProduct) await api.products.update(linkedProduct.id, { cost_price: finalUnitCost });
      }

      await onReload();
    } finally {
      setBusy(false);
    }
  };


  return (
    <div>
      <div className="card">
        <div className="card-title-row">
          <div className="card-title">Ishlab chiqarish — tannarxni hisoblash</div>
          <input className="input xs" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
        <table className="table" style={{ marginBottom: 10 }}>
          <thead><tr><th>Tayyor mahsulot</th><th className="right">Miqdor</th><th></th></tr></thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id}>
                <td>
                  <select className="input" value={r.finishedId} onChange={(e) => updateRow(r.id, { finishedId: e.target.value })}>
                    <option value="">— Tayyor mahsulotni tanlang —</option>
                    {finishedProducts.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
                  </select>
                </td>
                <td className="right"><input className="input mono sm" type="number" placeholder="Miqdor" value={r.qty} onChange={(e) => updateRow(r.id, { qty: e.target.value })} /></td>
                <td className="row-actions">{rows.length > 1 && <button className="icon-btn danger" onClick={() => removeRow(r.id)}><Trash2 size={14} /></button>}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="form-row wrap">
          <button className="btn btn-ghost" onClick={addRow}><Plus size={14} /> Yana mahsulot qo'shish</button>
          <button className="btn btn-ghost" onClick={calculate}><Settings2 size={14} /> Hisoblash</button>
        </div>

        {error && <div className="error-box"><AlertTriangle size={14} /> {error}</div>}

        {preview && (
          <div style={{ marginTop: 14 }}>
            {preview.map(({ row, finished, plan }, idx) => {
              const rowOverhead = overheadFor(finished);
              const rowOverheadTotal = rowOverhead * Number(row.qty);
              return (
                <div key={row.id} style={{ marginBottom: 16, paddingBottom: 14, borderBottom: idx < preview.length - 1 ? "1px dashed var(--border)" : "none" }}>
                  <div style={{ fontWeight: 600, marginBottom: 6 }}>{finished?.name} — {fmtQty(row.qty)} {finished?.unit} ({fmtQty(finished?.volume_liters || 1)} L/dona)</div>
                  <table className="table">
                    <thead><tr><th>Xom ashyo</th><th className="right">Kerak</th><th className="right">Tannarx</th></tr></thead>
                    <tbody>
                      {plan.details.map((d, i2) => (
                        <tr key={i2}>
                          <td>{d.material?.name}</td>
                          <td className="right mono">{fmtQty(d.required)} {d.material?.unit}</td>
                          <td className="right mono">{fmt(d.cost)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {plan.errors.length > 0 ? (
                    plan.errors.map((e, i3) => <div key={i3} className="error-box"><AlertTriangle size={14} /> {e}</div>)
                  ) : (
                    <div className="muted" style={{ fontSize: 12.5, marginTop: 6 }}>
                      Xom ashyo tannarxi: {fmt(plan.totalCost)} + Operatsion xarajat ({fmt(rowOverhead)}/L × {fmtQty(finished?.volume_liters || 1)} L × {row.qty}) = {fmt(rowOverheadTotal)}
                      <br />Jami: <b className="mono">{fmt(plan.totalCost + rowOverheadTotal)}</b> (birlik: {fmt(plan.unitCost + rowOverhead)})
                    </div>
                  )}
                </div>
              );
            })}
            {preview.every(({ plan }) => plan.errors.length === 0) && (
              <div className="invoice-footer">
                <div className="total-line">
                  Umumiy tannarx: <span className="mono total-amount">
                    {fmt(preview.reduce((s, { row, finished, plan }) => s + plan.totalCost + overheadFor(finished) * Number(row.qty), 0))}
                  </span>
                  <div className="muted" style={{ fontSize: 12 }}>Operatsion xarajat — oxirgi {overhead.days} kunlik xarajat ÷ shu davrda ishlab chiqarilgan jami litr asosida, har mahsulotning o'z litr hajmiga mutanosib taqsimlanadi.</div>
                </div>
                <button className="btn btn-primary" disabled={busy} onClick={confirmProduce}><Factory size={14} /> Tasdiqlash va kirim qilish</button>
              </div>
            )}
          </div>
        )}
      </div>

      <div className="card">
        <div className="card-title">Ishlab chiqarish tarixi ({productionBatches.length})</div>
        <table className="table">
          <thead><tr><th>Sana</th><th>Mahsulot</th><th className="right">Miqdor</th><th className="right">Birlik tannarx</th><th className="right">Jami tannarx</th><th></th></tr></thead>
          <tbody>
            {[...productionBatches].reverse().map((b) => (
              <tr key={b.id}>
                <td className="muted">{fmtDate(b.date)}</td>
                <td>{finishedProducts.find((f) => f.id === b.finished_product_id)?.name || "—"}</td>
                <td className="right mono">{b.qty}</td>
                <td className="right mono">{fmt(b.unit_cost)}</td>
                <td className="right mono">{fmt(b.total_cost)}</td>
                <td className="row-actions">
                  <button className="btn btn-ghost" style={{ padding: "4px 8px" }} disabled={busy} onClick={() => recalcProduction(b)}><Settings2 size={13} /> Qayta hisoblash</button>
                  <button className="icon-btn danger" disabled={busy} onClick={() => deleteProduction(b)}><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
            {productionBatches.length === 0 && <tr><td colSpan={6} className="empty">Hali ishlab chiqarilmagan</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------- Customers ---------------- */
function Customers({ customers, onReload }) {
  const [form, setForm] = useState({ name: "", phone: "", address: "", is_agent: false });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [busy, setBusy] = useState(false);

  const add = async () => {
    if (!form.name.trim()) return;
    setBusy(true);
    try {
      let commissionCustomerId = null;
      if (form.is_agent) {
        const commission = await api.customers.add({ name: `${form.name.trim()} — ish haqi`, phone: "", address: "" });
        commissionCustomerId = commission.id;
      }
      await api.customers.add({ name: form.name.trim(), phone: form.phone, address: form.address, is_agent: form.is_agent, commission_customer_id: commissionCustomerId });
      setForm({ name: "", phone: "", address: "", is_agent: false });
      await onReload();
    } finally { setBusy(false); }
  };
  const startEdit = (c) => { setEditingId(c.id); setEditForm({ ...c }); };
  const saveEdit = async () => {
    setBusy(true);
    try {
      let commissionCustomerId = editForm.commission_customer_id || null;
      if (editForm.is_agent && !commissionCustomerId) {
        const commission = await api.customers.add({ name: `${editForm.name} — ish haqi`, phone: "", address: "" });
        commissionCustomerId = commission.id;
      }
      await api.customers.update(editingId, {
        name: editForm.name, phone: editForm.phone, address: editForm.address,
        is_agent: editForm.is_agent, commission_customer_id: commissionCustomerId,
      });
      setEditingId(null);
      await onReload();
    } finally { setBusy(false); }
  };
  const remove = async (id) => {
    if (!confirm("Bu mijozni o'chirasizmi?")) return;
    setBusy(true);
    try { await api.customers.remove(id); await onReload(); } finally { setBusy(false); }
  };

  return (
    <div>
      <div className="card">
        <div className="card-title">Yangi mijoz qo'shish</div>
        <div className="form-row wrap">
          <input className="input" placeholder="Mijoz nomi" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
          <input className="input sm" placeholder="Telefon" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} />
          <input className="input" placeholder="Manzil (ixtiyoriy)" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} />
          <button className="btn btn-primary" disabled={busy} onClick={add}><Plus size={14} /> Qo'shish</button>
        </div>
        <label className="form-row" style={{ marginTop: 10, gap: 6, cursor: "pointer" }}>
          <input type="checkbox" checked={form.is_agent} onChange={(e) => setForm({ ...form, is_agent: e.target.checked })} />
          <span className="muted" style={{ fontSize: 13 }}>Bu mijoz — agent (vositachi). Faktura chiqarganda sotuv/kassa narxi farqi avtomatik "{form.name || "agent"} — ish haqi" hisobiga qarz sifatida yoziladi.</span>
        </label>
      </div>

      <div className="card">
        <div className="card-title">Mijozlar ro'yxati ({customers.length})</div>
        <table className="table">
          <thead><tr><th>Nomi</th><th>Telefon</th><th>Manzil</th><th>Agent</th><th></th></tr></thead>
          <tbody>
            {customers.map((c) => editingId === c.id ? (
              <tr key={c.id} className="editing-row">
                <td><input className="input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} /></td>
                <td><input className="input" value={editForm.phone || ""} onChange={(e) => setEditForm({ ...editForm, phone: e.target.value })} /></td>
                <td><input className="input" value={editForm.address || ""} onChange={(e) => setEditForm({ ...editForm, address: e.target.value })} /></td>
                <td><input type="checkbox" checked={!!editForm.is_agent} onChange={(e) => setEditForm({ ...editForm, is_agent: e.target.checked })} /></td>
                <td className="row-actions"><button className="icon-btn" disabled={busy} onClick={saveEdit}><Check size={15} /></button></td>
              </tr>
            ) : (
              <tr key={c.id}>
                <td>{c.name}</td>
                <td className="muted">{c.phone || "—"}</td>
                <td className="muted">{c.address || "—"}</td>
                <td className="muted">
                  {c.is_agent ? (
                    <span className="tone-ok-text" style={{ fontWeight: 600 }}>
                      Ha {c.commission_customer_id && `→ ${customers.find((x) => x.id === c.commission_customer_id)?.name || ""}`}
                    </span>
                  ) : "—"}
                </td>
                <td className="row-actions">
                  <button className="icon-btn" onClick={() => startEdit(c)}><Pencil size={14} /></button>
                  <button className="icon-btn danger" disabled={busy} onClick={() => remove(c.id)}><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
            {customers.length === 0 && <tr><td colSpan={5} className="empty">Hali mijoz qo'shilmagan</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------- Suppliers (Ta'minotchilar) ---------------- */
function Suppliers({ suppliers, supplierPayments, rawMaterialBatches, rawMaterials, onReload }) {
  const [form, setForm] = useState({ name: "", phone: "", address: "" });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [busy, setBusy] = useState(false);
  const [openId, setOpenId] = useState(null);
  const [payForm, setPayForm] = useState({ supplierId: "", amount: "", date: todayISO(), note: "" });

  const add = async () => {
    if (!form.name.trim()) return;
    setBusy(true);
    try { await api.suppliers.add(form); setForm({ name: "", phone: "", address: "" }); await onReload(); } finally { setBusy(false); }
  };
  const startEdit = (s) => { setEditingId(s.id); setEditForm({ ...s }); };
  const saveEdit = async () => {
    setBusy(true);
    try { await api.suppliers.update(editingId, { name: editForm.name, phone: editForm.phone, address: editForm.address }); setEditingId(null); await onReload(); } finally { setBusy(false); }
  };
  const remove = async (id) => {
    if (!confirm("Bu ta'minotchini o'chirasizmi?")) return;
    setBusy(true);
    try { await api.suppliers.remove(id); await onReload(); } finally { setBusy(false); }
  };

  const addPayment = async () => {
    if (!payForm.supplierId || !Number(payForm.amount)) return;
    setBusy(true);
    try {
      await api.supplierPayments.add({ supplier_id: payForm.supplierId, amount: Number(payForm.amount), date: payForm.date, note: payForm.note });
      setPayForm({ supplierId: "", amount: "", date: todayISO(), note: "" });
      await onReload();
    } finally { setBusy(false); }
  };
  const removePayment = async (id) => {
    if (!confirm("Bu to'lovni o'chirasizmi?")) return;
    await api.supplierPayments.remove(id);
    await onReload();
  };

  const ledger = (supplierId) => {
    const purchased = rawMaterialBatches.filter((b) => b.supplier_id === supplierId).reduce((s, b) => s + Number(b.qty) * Number(b.unit_cost), 0);
    const paid = supplierPayments.filter((p) => p.supplier_id === supplierId).reduce((s, p) => s + Number(p.amount), 0);
    return { purchased, paid, balance: purchased - paid };
  };

  return (
    <div>
      <div className="card">
        <div className="card-title">Yangi ta'minotchi qo'shish</div>
        <div className="form-row wrap">
          <input className="input" placeholder="Ta'minotchi nomi" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
          <input className="input sm" placeholder="Telefon" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} />
          <input className="input" placeholder="Manzil (ixtiyoriy)" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} />
          <button className="btn btn-primary" disabled={busy} onClick={add}><Plus size={14} /> Qo'shish</button>
        </div>
      </div>

      <div className="card">
        <div className="card-title">Ta'minotchi(lar)ga to'lov qilish</div>
        <div className="form-row wrap">
          <select className="input" value={payForm.supplierId} onChange={(e) => setPayForm({ ...payForm, supplierId: e.target.value })}>
            <option value="">— Ta'minotchini tanlang —</option>
            {suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
          <input className="input sm mono" type="number" placeholder="Summa" value={payForm.amount} onChange={(e) => setPayForm({ ...payForm, amount: e.target.value })} />
          <input className="input xs" type="date" value={payForm.date} onChange={(e) => setPayForm({ ...payForm, date: e.target.value })} />
          <input className="input" placeholder="Izoh (ixtiyoriy)" value={payForm.note} onChange={(e) => setPayForm({ ...payForm, note: e.target.value })} />
          <button className="btn btn-primary" disabled={busy} onClick={addPayment}><Plus size={14} /> To'lov qo'shish</button>
        </div>
      </div>

      <div className="card">
        <div className="card-title">Ta'minotchilar bilan hisob-kitob ({suppliers.length})</div>
        <table className="table">
          <thead><tr><th>Ta'minotchi</th><th className="right">Jami xarid</th><th className="right">To'langan</th><th className="right">Bizning qarzimiz</th><th></th><th></th></tr></thead>
          <tbody>
            {suppliers.map((s) => {
              const l = ledger(s.id);
              const isOpen = openId === s.id;
              const isEditing = editingId === s.id;
              const batchHistory = rawMaterialBatches.filter((b) => b.supplier_id === s.id).map((b) => ({
                date: b.date, label: `${rawMaterials.find((m) => m.id === b.raw_material_id)?.name || "?"} — ${fmtQty(b.qty)} × ${fmt(b.unit_cost)}`,
                amount: Number(b.qty) * Number(b.unit_cost), type: "purchase",
              }));
              const paymentHistory = supplierPayments.filter((p) => p.supplier_id === s.id).map((p) => ({
                id: p.id, date: p.date, label: p.note ? `To'lov — ${p.note}` : "To'lov", amount: -Number(p.amount), type: "payment",
              }));
              const history = [...batchHistory, ...paymentHistory].sort((a, b) => new Date(a.date) - new Date(b.date));
              return isEditing ? (
                <tr key={s.id} className="editing-row">
                  <td><input className="input" value={editForm.name} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} /></td>
                  <td colSpan={3}>
                    <div className="form-row wrap">
                      <input className="input sm" placeholder="Telefon" value={editForm.phone || ""} onChange={(e) => setEditForm({ ...editForm, phone: e.target.value })} />
                      <input className="input" placeholder="Manzil" value={editForm.address || ""} onChange={(e) => setEditForm({ ...editForm, address: e.target.value })} />
                    </div>
                  </td>
                  <td className="row-actions"><button className="icon-btn" disabled={busy} onClick={saveEdit}><Check size={15} /></button></td>
                  <td></td>
                </tr>
              ) : (
                <React.Fragment key={s.id}>
                  <tr className="clickable-row">
                    <td onClick={() => setOpenId(isOpen ? null : s.id)}>{s.name} <span className="muted">{s.phone ? `· ${s.phone}` : ""}</span></td>
                    <td className="right mono" onClick={() => setOpenId(isOpen ? null : s.id)}>{fmt(l.purchased)}</td>
                    <td className="right mono tone-ok-text" onClick={() => setOpenId(isOpen ? null : s.id)}>{fmt(l.paid)}</td>
                    <td className={`right mono ${l.balance > 0 ? "tone-debt-text" : ""}`} onClick={() => setOpenId(isOpen ? null : s.id)}>{fmt(l.balance)}</td>
                    <td className="row-actions">
                      <button className="icon-btn" onClick={() => startEdit(s)}><Pencil size={14} /></button>
                      <button className="icon-btn danger" onClick={() => remove(s.id)}><Trash2 size={14} /></button>
                    </td>
                    <td className="muted" onClick={() => setOpenId(isOpen ? null : s.id)}>{isOpen ? "▲" : "▼"}</td>
                  </tr>
                  {isOpen && (
                    <tr>
                      <td colSpan={6} className="history-cell">
                        {history.length === 0 ? <div className="empty">Tarix yo'q</div> : (
                          <table className="table sub-table">
                            <tbody>
                              {history.map((h, idx) => (
                                <tr key={idx}>
                                  <td className="muted">{fmtDate(h.date)}</td>
                                  <td>{h.label}</td>
                                  <td className={`right mono ${h.amount < 0 ? "tone-ok-text" : ""}`}>{h.amount < 0 ? `+${fmt(-h.amount)}` : fmt(h.amount)}</td>
                                  <td className="row-actions">
                                    {h.type === "payment" && <button className="icon-btn danger" onClick={() => removePayment(h.id)}><Trash2 size={14} /></button>}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        )}
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
            {suppliers.length === 0 && <tr><td colSpan={6} className="empty">Hali ta'minotchi qo'shilmagan</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------- Invoice ---------------- */
function InvoiceTab({ products, customers, invoices, currentUser, onReload, onPrint, onNeedProduction, customerId, setCustomerId, date, setDate, rows, setRows }) {
  const [error, setError] = useState("");
  const [stockShortage, setStockShortage] = useState(null); // { productId, name, shortfall }
  const [busy, setBusy] = useState(false);
  const [editingInvoiceId, setEditingInvoiceId] = useState(null);

  const customer = customers.find((c) => c.id === customerId);
  const isAgentInvoice = !!customer?.is_agent;

  const addRow = () => setRows([...rows, { id: Math.random().toString(36).slice(2), product_id: "", qty: 1, price: 0, base_price: 0 }]);
  const updateRow = (id, patch) => setRows(rows.map((r) => {
    if (r.id !== id) return r;
    const next = { ...r, ...patch };
    if (patch.product_id) {
      const p = products.find((x) => x.id === patch.product_id);
      if (p) { next.price = p.price; next.base_price = p.price; }
    }
    return next;
  }));
  const removeRow = (id) => setRows(rows.filter((r) => r.id !== id));

  const total = rows.reduce((s, r) => s + Number(r.qty) * Number(r.price), 0);
  const commission = isAgentInvoice ? rows.reduce((s, r) => s + (Number(r.price) - Number(r.base_price)) * Number(r.qty), 0) : 0;

  const startEditInvoice = (inv) => {
    setEditingInvoiceId(inv.id);
    setCustomerId(inv.customer_id);
    setDate(inv.date);
    setRows((inv.items || []).filter((it) => it.product_id).map((it) => ({
      id: Math.random().toString(36).slice(2), product_id: it.product_id, qty: it.qty, price: it.price, base_price: it.price,
    })));
    setError("");
  };
  const cancelEdit = () => { setEditingInvoiceId(null); setRows([]); setCustomerId(""); setError(""); };

  const save = async () => {
    setError(""); setStockShortage(null);
    if (!customerId) { setError("Mijozni tanlang"); return; }
    if (rows.length === 0) { setError("Kamida bitta mahsulot qo'shing"); return; }
    for (const r of rows) {
      const p = products.find((x) => x.id === r.product_id);
      if (!p) { setError("Barcha qatorlarda mahsulot tanlanishi kerak"); return; }
      if (Number(r.qty) <= 0) { setError(`"${p.name}" uchun miqdor noto'g'ri`); return; }
    }
    if (isAgentInvoice && !customer.commission_customer_id) { setError("Bu agent uchun ish haqi hisobi bog'lanmagan (Mijozlar bo'limida tekshiring)"); return; }

    const oldInvoice = editingInvoiceId ? invoices.find((i) => i.id === editingInvoiceId) : null;
    const oldQtyMap = {};
    if (oldInvoice) (oldInvoice.items || []).forEach((it) => { if (it.product_id) oldQtyMap[it.product_id] = (oldQtyMap[it.product_id] || 0) + Number(it.qty); });
    const newQtyMap = {};
    rows.forEach((r) => { newQtyMap[r.product_id] = (newQtyMap[r.product_id] || 0) + Number(r.qty); });

    for (const pid of Object.keys(newQtyMap)) {
      const p = products.find((x) => x.id === pid);
      const available = Number(p.qty) + (oldQtyMap[pid] || 0);
      if (newQtyMap[pid] > available) {
        setError(`"${p.name}" uchun omborda yetarli qoldiq yo'q (bor: ${available})`);
        setStockShortage({ productId: pid, name: p.name, shortfall: newQtyMap[pid] - available });
        return;
      }
    }

    setBusy(true);
    try {
      const items = rows.map((r) => {
        const p = products.find((x) => x.id === r.product_id);
        return { product_id: p.id, name: p.name, unit: p.unit, qty: Number(r.qty), price: Number(r.price) };
      });

      const allProductIds = new Set([...Object.keys(oldQtyMap), ...Object.keys(newQtyMap)]);
      for (const pid of allProductIds) {
        const p = products.find((x) => x.id === pid);
        if (!p) continue;
        const delta = (oldQtyMap[pid] || 0) - (newQtyMap[pid] || 0);
        if (delta === 0) continue;
        await api.products.update(pid, { qty: Number(p.qty) + delta });
        await api.stockMovements.add({
          product_id: pid, date, qty: delta, type: delta > 0 ? "tuzatish" : "chiqim",
          note: oldInvoice ? `Faktura ${oldInvoice.number} tahrirlandi` : `Faktura`,
        });
      }

      let invoiceNumber, invoiceRecord;
      if (oldInvoice) {
        invoiceNumber = oldInvoice.number;
        invoiceRecord = await api.invoices.update(oldInvoice.id, { date, customer_id: customerId, items, total });
        const oldCommission = invoices.find((i) => i.number === `${oldInvoice.number}-K`);
        if (oldCommission) await api.invoices.remove(oldCommission.id);
      } else {
        invoiceNumber = `F-${String(invoices.length + 1).padStart(4, "0")}`;
        invoiceRecord = await api.invoices.add({ number: invoiceNumber, date, customer_id: customerId, items, total, created_by: currentUser || null });
      }

      if (isAgentInvoice && commission !== 0 && customer.commission_customer_id) {
        await api.invoices.add({
          number: `${invoiceNumber}-K`, date, customer_id: customer.commission_customer_id,
          items: [{ name: `Agent komissiyasi — Faktura ${invoiceNumber}`, unit: "", qty: 1, price: commission }],
          total: commission, created_by: currentUser || null,
        });
      }

      setRows([]); setCustomerId(""); setEditingInvoiceId(null);
      await onReload();
      onPrint(invoiceRecord);
    } catch (e) {
      setError(e.message || "Saqlashda xatolik yuz berdi");
    } finally {
      setBusy(false);
    }
  };

  const removeInvoice = async (id) => {
    if (!confirm("Bu fakturani o'chirasizmi? (Ombor qoldig'i qaytarilmaydi)")) return;
    await api.invoices.remove(id);
    await onReload();
  };

  return (
    <div>
      <div className="card">
        <div className="card-title-row">
          <div className="card-title">{editingInvoiceId ? "Fakturani tahrirlash" : "Yangi faktura"}</div>
          {editingInvoiceId && <button className="btn btn-ghost" onClick={cancelEdit}><X size={14} /> Bekor qilish</button>}
        </div>
        <div className="form-row wrap">
          <select className="input" value={customerId} onChange={(e) => setCustomerId(e.target.value)}>
            <option value="">— Mijozni tanlang —</option>
            {customers.filter((c) => !customers.some((x) => x.commission_customer_id === c.id)).map((c) => <option key={c.id} value={c.id}>{c.name}{c.is_agent ? " (agent)" : ""}</option>)}
          </select>
          <input className="input xs" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
        {isAgentInvoice && (
          <div className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>
            Bu — agent. Har bir qator uchun "Sotuv narxi" va "Kassa narxi"ni kiriting; farqi × miqdor avtomatik <b>{customer?.name} — ish haqi</b> hisobiga qarz sifatida yoziladi.
          </div>
        )}

        <table className="table" style={{ marginTop: 12 }}>
          <thead>
            <tr>
              <th>Mahsulot</th><th className="right">Miqdor</th>
              <th className="right">{isAgentInvoice ? "Sotuv narxi" : "Narxi"}</th>
              {isAgentInvoice && <th className="right">Kassa narxi</th>}
              <th className="right">Summa</th><th></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id}>
                <td>
                  <select className="input" value={r.product_id} onChange={(e) => updateRow(r.id, { product_id: e.target.value })}>
                    <option value="">— tanlang —</option>
                    {products.map((p) => <option key={p.id} value={p.id}>{p.name} (bor: {p.qty})</option>)}
                  </select>
                </td>
                <td className="right"><input className="input mono xs" type="number" value={r.qty} onChange={(e) => updateRow(r.id, { qty: e.target.value })} /></td>
                <td className="right"><input className="input mono sm" type="number" value={r.price} onChange={(e) => updateRow(r.id, { price: e.target.value })} /></td>
                {isAgentInvoice && (
                  <td className="right"><input className="input mono sm" type="number" value={r.base_price} onChange={(e) => updateRow(r.id, { base_price: e.target.value })} /></td>
                )}
                <td className="right mono">{fmt(Number(r.qty) * Number(r.price))}</td>
                <td className="row-actions"><button className="icon-btn danger" onClick={() => removeRow(r.id)}><Trash2 size={14} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>

        <button className="btn btn-ghost" onClick={addRow}><Plus size={14} /> Qator qo'shish</button>

        {error && (
          <div className="error-box" style={{ justifyContent: "space-between" }}>
            <span><AlertTriangle size={14} /> {error}</span>
            {stockShortage && (
              <button
                className="btn btn-primary" style={{ padding: "4px 10px", fontSize: 12 }}
                onClick={() => onNeedProduction(stockShortage.productId, stockShortage.shortfall)}
              >
                <Factory size={13} /> Ishlab chiqarishga o'tish
              </button>
            )}
          </div>
        )}

        <div className="invoice-footer">
          <div className="total-line">
            Jami: <span className="mono total-amount">{fmt(total)}</span>
            {isAgentInvoice && <div className="muted" style={{ fontSize: 12.5 }}>Agent ish haqi: {fmt(commission)}</div>}
          </div>
          <button className="btn btn-primary" disabled={busy} onClick={save}><FileText size={14} /> {editingInvoiceId ? "Yangilash" : "Saqlash va chop etish"}</button>
        </div>
      </div>

      <div className="card">
        <div className="card-title">Chiqarilgan fakturalar ({invoices.length})</div>
        <table className="table">
          <thead><tr><th>№</th><th>Sana</th><th>Mijoz</th><th className="right">Summa</th><th>Kim kiritdi</th><th></th></tr></thead>
          <tbody>
            {[...invoices].reverse().map((inv) => (
              <tr key={inv.id}>
                <td className="mono">{inv.number}</td>
                <td className="muted">{fmtDate(inv.date)}</td>
                <td>{customers.find((c) => c.id === inv.customer_id)?.name || "—"}</td>
                <td className="right mono">{fmt(inv.total)}</td>
                <td className="muted">{inv.created_by || "—"}</td>
                <td className="row-actions">
                  <button className="icon-btn" onClick={() => onPrint(inv)}><Printer size={14} /></button>
                  {!inv.number.endsWith("-K") && <button className="icon-btn" onClick={() => startEditInvoice(inv)}><Pencil size={14} /></button>}
                  <button className="icon-btn danger" onClick={() => removeInvoice(inv.id)}><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
            {invoices.length === 0 && <tr><td colSpan={6} className="empty">Hali faktura chiqarilmagan</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------- Payments ---------------- */
function Payments({ customers, payments, onReload }) {
  const [form, setForm] = useState({ customer_id: "", amount: "", date: todayISO(), note: "" });
  const [busy, setBusy] = useState(false);

  const add = async () => {
    if (!form.customer_id || !Number(form.amount)) return;
    setBusy(true);
    try {
      await api.payments.add({ customer_id: form.customer_id, amount: Number(form.amount), date: form.date, note: form.note });
      setForm({ customer_id: "", amount: "", date: todayISO(), note: "" });
      await onReload();
    } finally { setBusy(false); }
  };
  const remove = async (id) => {
    if (!confirm("Bu to'lov yozuvini o'chirasizmi?")) return;
    await api.payments.remove(id);
    await onReload();
  };

  return (
    <div>
      <div className="card">
        <div className="card-title">Tushum qo'shish</div>
        <div className="form-row wrap">
          <select className="input" value={form.customer_id} onChange={(e) => setForm({ ...form, customer_id: e.target.value })}>
            <option value="">— Mijozni tanlang —</option>
            {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          <input className="input sm mono" type="number" placeholder="Summa" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} />
          <input className="input xs" type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
          <input className="input" placeholder="Izoh (ixtiyoriy)" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
          <button className="btn btn-primary" disabled={busy} onClick={add}><Plus size={14} /> Qo'shish</button>
        </div>
      </div>

      <div className="card">
        <div className="card-title">To'lovlar tarixi ({payments.length})</div>
        <table className="table">
          <thead><tr><th>Sana</th><th>Mijoz</th><th className="right">Summa</th><th>Izoh</th><th></th></tr></thead>
          <tbody>
            {[...payments].reverse().map((p) => (
              <tr key={p.id}>
                <td className="muted">{fmtDate(p.date)}</td>
                <td>{customers.find((c) => c.id === p.customer_id)?.name || "—"}</td>
                <td className="right mono tone-ok-text">{fmt(p.amount)}</td>
                <td className="muted">{p.note || "—"}</td>
                <td className="row-actions"><button className="icon-btn danger" onClick={() => remove(p.id)}><Trash2 size={14} /></button></td>
              </tr>
            ))}
            {payments.length === 0 && <tr><td colSpan={5} className="empty">Hali to'lov qayd etilmagan</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------- Ledger ---------------- */
function buildAct(customerId, invoices, payments, from, to) {
  const custInvoices = invoices.filter((i) => i.customer_id === customerId);
  const custPayments = payments.filter((p) => p.customer_id === customerId);

  let opening = 0;
  if (from) {
    opening =
      custInvoices.filter((i) => i.date < from).reduce((s, i) => s + Number(i.total), 0) -
      custPayments.filter((p) => p.date < from).reduce((s, p) => s + Number(p.amount), 0);
  }

  const rows = [
    ...custInvoices
      .filter((i) => (!from || i.date >= from) && (!to || i.date <= to))
      .map((i) => ({ date: i.date, label: `Faktura ${i.number}`, debit: Number(i.total), credit: 0 })),
    ...custPayments
      .filter((p) => (!from || p.date >= from) && (!to || p.date <= to))
      .map((p) => ({ date: p.date, label: p.note ? `To'lov — ${p.note}` : "To'lov", debit: 0, credit: Number(p.amount) })),
  ].sort((a, b) => new Date(a.date) - new Date(b.date));

  let running = opening;
  const withBalance = rows.map((r) => { running += r.debit - r.credit; return { ...r, balance: running }; });
  return { opening, rows: withBalance, closing: running };
}

function Ledger({ customers, invoices, payments, balances, settings, onReload }) {
  const [openId, setOpenId] = useState(null);
  const [actModal, setActModal] = useState(null); // { customer, from, to }
  const [actPrint, setActPrint] = useState(null);
  const [reportFrom, setReportFrom] = useState("");
  const [reportTo, setReportTo] = useState(todayISO());

  const openActModal = (customer) => setActModal({ customer, from: "", to: todayISO() });

  const generateAct = () => {
    const { customer, from, to } = actModal;
    const data = buildAct(customer.id, invoices, payments, from, to);
    setActPrint({ customer, from, to, ...data });
    setActModal(null);
  };

  const removePayment = async (id) => {
    if (!confirm("Bu to'lov yozuvini o'chirasizmi? (Agar Kassa'da ham yozuv bo'lsa, uni alohida o'chirish kerak)")) return;
    await api.payments.remove(id);
    await onReload();
  };

  const rangeData = (customerId) => {
    const custInvoices = invoices.filter((i) => i.customer_id === customerId);
    const custPayments = payments.filter((p) => p.customer_id === customerId);
    const opening = reportFrom
      ? custInvoices.filter((i) => i.date < reportFrom).reduce((s, i) => s + Number(i.total), 0) - custPayments.filter((p) => p.date < reportFrom).reduce((s, p) => s + Number(p.amount), 0)
      : 0;
    const kirim = custInvoices.filter((i) => (!reportFrom || i.date >= reportFrom) && (!reportTo || i.date <= reportTo)).reduce((s, i) => s + Number(i.total), 0);
    const chiqim = custPayments.filter((p) => (!reportFrom || p.date >= reportFrom) && (!reportTo || p.date <= reportTo)).reduce((s, p) => s + Number(p.amount), 0);
    const closing = opening + kirim - chiqim;
    return { opening, kirim, chiqim, closing };
  };

  return (
    <div className="card" style={{ fontSize: 12.5 }}>
      <div className="card-title">Mijozlar bilan hisob-kitob (vzaimoraschet)</div>
      <div className="form-row wrap" style={{ marginBottom: 14 }}>
        <label className="muted" style={{ fontSize: 12 }}>Dan:</label>
        <input className="input xs" type="date" value={reportFrom} onChange={(e) => setReportFrom(e.target.value)} />
        <label className="muted" style={{ fontSize: 12 }}>Gacha:</label>
        <input className="input xs" type="date" value={reportTo} onChange={(e) => setReportTo(e.target.value)} />
        {reportFrom && <button className="btn btn-ghost" onClick={() => setReportFrom("")}>Boshidan ko'rsatish</button>}
      </div>
      <table className="table" style={{ fontSize: 12.5 }}>
        <thead><tr><th>Mijoz</th><th className="right">Boshiga qoldiq</th><th className="right">Kirim (faktura)</th><th className="right">Chiqim (to'lov)</th><th className="right">Oxiriga qoldiq</th><th></th><th></th></tr></thead>
        <tbody>
          {customers.map((c) => {
            const r = rangeData(c.id);
            const isOpen = openId === c.id;
            const history = [
              ...invoices.filter((i) => i.customer_id === c.id).map((i) => ({ date: i.date, label: `Faktura ${i.number}`, amount: Number(i.total), type: "invoice" })),
              ...payments.filter((p) => p.customer_id === c.id).map((p) => ({ id: p.id, date: p.date, label: p.note ? `To'lov — ${p.note}` : "To'lov", amount: -Number(p.amount), type: "payment" })),
            ].sort((a, b2) => new Date(a.date) - new Date(b2.date));
            return (
              <React.Fragment key={c.id}>
                <tr className="clickable-row">
                  <td onClick={() => setOpenId(isOpen ? null : c.id)}>{c.name}</td>
                  <td className="right mono" onClick={() => setOpenId(isOpen ? null : c.id)}>{fmtNum(r.opening)}</td>
                  <td className="right mono tone-debt-text" onClick={() => setOpenId(isOpen ? null : c.id)}>{fmtNum(r.kirim)}</td>
                  <td className="right mono tone-ok-text" onClick={() => setOpenId(isOpen ? null : c.id)}>{fmtNum(r.chiqim)}</td>
                  <td className={`right mono ${r.closing > 0 ? "tone-debt-text" : r.closing < 0 ? "tone-ok-text" : ""}`} onClick={() => setOpenId(isOpen ? null : c.id)}>
                    {r.closing > 0 ? `Qarz: ${fmtNum(r.closing)}` : r.closing < 0 ? `Ortiqcha: ${fmtNum(-r.closing)}` : "Tenglashgan"}
                  </td>
                  <td>
                    <button className="btn btn-ghost" style={{ padding: "4px 8px", fontSize: 12 }} onClick={() => openActModal(c)}><ClipboardList size={13} /> Akt sverka</button>
                  </td>
                  <td className="muted" onClick={() => setOpenId(isOpen ? null : c.id)}>{isOpen ? "▲" : "▼"}</td>
                </tr>
                {isOpen && (
                  <tr>
                    <td colSpan={7} className="history-cell">
                      {history.length === 0 ? <div className="empty">Tarix yo'q</div> : (
                        <table className="table sub-table">
                          <tbody>
                            {history.map((h, idx) => (
                              <tr key={idx}>
                                <td className="muted">{fmtDate(h.date)}</td>
                                <td>{h.label}</td>
                                <td className={`right mono ${h.amount < 0 ? "tone-ok-text" : ""}`}>{h.amount < 0 ? `+${fmtNum(-h.amount)}` : fmtNum(h.amount)}</td>
                                <td className="row-actions">
                                  {h.type === "payment" && <button className="icon-btn danger" onClick={() => removePayment(h.id)}><Trash2 size={14} /></button>}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </td>
                  </tr>
                )}
              </React.Fragment>
            );
          })}
          {customers.length === 0 && <tr><td colSpan={7} className="empty">Hali mijoz yo'q</td></tr>}
        </tbody>
      </table>

      {actModal && (
        <div className="modal-backdrop" onClick={() => setActModal(null)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="card-title">Akt sverka — {actModal.customer.name}</div>
            <div className="muted" style={{ marginBottom: 10, fontSize: 12.5 }}>
              Davrni tanlang (bo'sh qoldirsangiz — butun davr bo'yicha chiqadi)
            </div>
            <div className="form-row wrap">
              <input className="input" type="date" value={actModal.from} onChange={(e) => setActModal({ ...actModal, from: e.target.value })} />
              <input className="input" type="date" value={actModal.to} onChange={(e) => setActModal({ ...actModal, to: e.target.value })} />
            </div>
            <div className="invoice-footer" style={{ marginTop: 14 }}>
              <button className="btn btn-ghost" onClick={() => setActModal(null)}>Bekor qilish</button>
              <button className="btn btn-primary" onClick={generateAct}><ClipboardList size={14} /> Shakllantirish</button>
            </div>
          </div>
        </div>
      )}

      {actPrint && <ActPrintOverlay act={actPrint} settings={settings} onClose={() => setActPrint(null)} />}
    </div>
  );
}

/* ---------------- Akt sverka chop etish ---------------- */
function ActPrintOverlay({ act, settings, onClose }) {
  const { customer, rows, opening, closing, from, to } = act;
  return (
    <div className="print-backdrop">
      <div className="print-toolbar no-print">
        <button className="btn btn-ghost" onClick={onClose}><X size={14} /> Yopish</button>
        <button className="btn btn-primary" onClick={() => window.print()}><Printer size={14} /> Chop etish / PDF saqlash</button>
      </div>
      <div id="print-area" className="invoice-sheet">
        <div className="invoice-head">
          <div>
            <div className="invoice-company">{settings.company_name}</div>
            {settings.company_phone && <div className="muted">{settings.company_phone}</div>}
          </div>
          <div className="invoice-meta">
            <div><b>Akt sverka</b></div>
            <div>{from ? `${fmtDate(from)} — ${fmtDate(to || todayISO())}` : "Butun davr bo'yicha"}</div>
          </div>
        </div>
        <div className="invoice-to"><b>Mijoz:</b> {customer.name} {customer.phone ? `· ${customer.phone}` : ""}</div>
        <table className="table invoice-table">
          <thead><tr><th>Sana</th><th>Hujjat</th><th className="right">Debet (faktura)</th><th className="right">Kredit (to'lov)</th><th className="right">Qoldiq</th></tr></thead>
          <tbody>
            {from && (
              <tr>
                <td colSpan={4}><i>Boshlang'ich qoldiq</i></td>
                <td className="right mono">{fmt(opening)}</td>
              </tr>
            )}
            {rows.map((r, idx) => (
              <tr key={idx}>
                <td className="muted">{fmtDate(r.date)}</td>
                <td>{r.label}</td>
                <td className="right mono">{r.debit ? fmt(r.debit) : "—"}</td>
                <td className="right mono">{r.credit ? fmt(r.credit) : "—"}</td>
                <td className="right mono">{fmt(r.balance)}</td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={5} className="empty">Bu davrda operatsiya yo'q</td></tr>}
          </tbody>
        </table>
        <div className="invoice-total">
          Yakuniy qoldiq: <span className="mono">{fmt(Math.abs(closing))}</span>
          {closing > 0 ? " (mijoz qarzdor)" : closing < 0 ? " (ortiqcha to'langan)" : " (tenglashgan)"}
        </div>
        <div className="invoice-sign-row">
          <div>Yetkazib beruvchi: ______________</div>
          <div>Xaridor: ______________</div>
        </div>
      </div>
    </div>
  );
}

/* ---------------- Print overlay ---------------- */
function PrintOverlay({ invoice, customers, settings, onClose }) {
  const customer = customers.find((c) => c.id === invoice.customer_id);
  return (
    <div className="print-backdrop">
      <div className="print-toolbar no-print">
        <button className="btn btn-ghost" onClick={onClose}><X size={14} /> Yopish</button>
        <button className="btn btn-primary" onClick={() => window.print()}><Printer size={14} /> Chop etish / PDF saqlash</button>
      </div>
      <div id="print-area" className="invoice-sheet">
        <div className="invoice-head">
          <div>
            <div className="invoice-company">{settings.company_name}</div>
            {settings.company_phone && <div className="muted">{settings.company_phone}</div>}
          </div>
          <div className="invoice-meta">
            <div><b>Faktura №</b> {invoice.number}</div>
            <div><b>Sana:</b> {fmtDate(invoice.date)}</div>
          </div>
        </div>
        <div className="invoice-to"><b>Mijoz:</b> {customer?.name || "—"} {customer?.phone ? `· ${customer.phone}` : ""}</div>
        <table className="table invoice-table">
          <thead><tr><th>№</th><th>Mahsulot</th><th className="right">Miqdor</th><th className="right">Narxi</th><th className="right">Summa</th></tr></thead>
          <tbody>
            {invoice.items.map((it, idx) => (
              <tr key={idx}>
                <td>{idx + 1}</td>
                <td>{it.name}</td>
                <td className="right mono">{it.qty} {it.unit}</td>
                <td className="right mono">{fmt(it.price)}</td>
                <td className="right mono">{fmt(it.qty * it.price)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="invoice-total">Jami to'lov: <span className="mono">{fmt(invoice.total)}</span></div>
        <div className="invoice-sign-row">
          <div>Topshirdi: ______________</div>
          <div>Qabul qildi: ______________</div>
        </div>
      </div>
    </div>
  );
}

/* ---------------- Analiz ---------------- */
function avgMaterialCost(materialId, batches) {
  const active = batches.filter((b) => b.raw_material_id === materialId && b.active && Number(b.remaining_qty) > 0);
  if (active.length > 0) {
    const totalQty = active.reduce((s, b) => s + Number(b.remaining_qty), 0);
    const totalVal = active.reduce((s, b) => s + Number(b.remaining_qty) * Number(b.unit_cost), 0);
    return totalQty > 0 ? totalVal / totalQty : 0;
  }
  const all = batches.filter((b) => b.raw_material_id === materialId).sort((a, b) => new Date(b.date) - new Date(a.date));
  return all.length > 0 ? Number(all[0].unit_cost) : 0;
}

function Analytics({ products, customers, invoices, payments, cashTransactions, balances, finishedProducts, productNorms, rawMaterials, rawMaterialBatches, productionBatches, fixedAssets }) {
  const [report, setReport] = useState(null);
  const [from, setFrom] = useState("");
  const [to, setTo] = useState(todayISO());
  const [pcProductId, setPcProductId] = useState("");
  const [pcPrice, setPcPrice] = useState("");
  const [pcMaterialPrices, setPcMaterialPrices] = useState({});

  const realInvoices = invoices.filter((i) => !i.number.endsWith("-K"));
  const inRange = (d) => (!from || d >= from) && (!to || d <= to);

  const REPORTS = [
    { id: "profitability", label: "Mahsulot foydaliligi", icon: BarChart3 },
    { id: "pl", label: "Profit & Loss", icon: FileText },
    { id: "customers", label: "Mijozlar kesimida tahlil", icon: Users },
    { id: "topProducts", label: "Eng ko'p sotilgan mahsulotlar", icon: Package },
    { id: "rfm", label: "Mijozlar segmentatsiyasi (RFM)", icon: TrendingUp },
    { id: "seasonal", label: "Oylar bo'yicha dinamika", icon: TrendingUp },
    { id: "agentRating", label: "Agentlar reytingi", icon: Award },
    { id: "pricingCalc", label: "Narxlash kalkulyatori", icon: Calculator },
    { id: "debtAging", label: "Qarzdorlik muddati", icon: AlertTriangle },
  ];

  return (
    <div>
      <div className="card">
        <div className="card-title">Analiz turini tanlang</div>
        <div className="form-row wrap">
          {REPORTS.map((r) => (
            <button key={r.id} className={`btn ${report === r.id ? "btn-primary" : "btn-ghost"}`} onClick={() => setReport(r.id)}>
              <r.icon size={14} /> {r.label}
            </button>
          ))}
        </div>
      </div>

      {report && (report === "pl" || report === "customers" || report === "topProducts") && (
        <div className="card">
          <div className="form-row wrap">
            <label className="muted" style={{ fontSize: 12.5 }}>Dan:</label>
            <input className="input xs" type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
            <label className="muted" style={{ fontSize: 12.5 }}>Gacha:</label>
            <input className="input xs" type="date" value={to} onChange={(e) => setTo(e.target.value)} />
            {from && <button className="btn btn-ghost" onClick={() => setFrom("")}>Boshidan</button>}
          </div>
        </div>
      )}

      {report === "profitability" && (() => {
        const sortedInvoicesDesc = [...realInvoices].sort((a, b) => new Date(b.date) - new Date(a.date));
        return (
          <div className="card">
            <div className="card-title">Mahsulot foydaliligi (1 birlik uchun)</div>
            <div className="muted" style={{ fontSize: 12, marginBottom: 8 }}>Sotish narxi — oxirgi 10 ta fakturadagi shu mahsulotning miqdorga tortilgan o'rtacha narxi</div>
            <table className="table">
              <thead><tr><th>Mahsulot</th><th className="right">O'rtacha sotish narxi</th><th className="right">Tannarx</th><th className="right">Foyda (birlik)</th><th className="right">Foyda %</th></tr></thead>
              <tbody>
                {products.map((p) => {
                  const cost = Number(p.cost_price) || 0;
                  const recentAvg = (() => {
                    const entries = [];
                    for (const inv of sortedInvoicesDesc) {
                      const item = (inv.items || []).find((it) => it.product_id === p.id);
                      if (item) entries.push(item);
                      if (entries.length >= 10) break;
                    }
                    if (entries.length === 0) return null;
                    const totalQty = entries.reduce((s, e) => s + Number(e.qty), 0);
                    const totalValue = entries.reduce((s, e) => s + Number(e.qty) * Number(e.price), 0);
                    return totalQty > 0 ? totalValue / totalQty : null;
                  })();
                  const price = recentAvg !== null ? recentAvg : Number(p.price) || 0;
                  const margin = price - cost;
                  const marginPct = price > 0 ? (margin / price) * 100 : 0;
                  return (
                    <tr key={p.id}>
                      <td>{p.name} <span className="muted">({p.unit})</span></td>
                      <td className="right mono">
                        {fmt(price)}
                        {recentAvg === null && <div className="muted" style={{ fontSize: 11 }}>faktura tarixi yo'q, mahsulot narxi</div>}
                      </td>
                      <td className="right mono">{fmt(cost)}</td>
                      <td className={`right mono ${margin >= 0 ? "tone-ok-text" : "tone-debt-text"}`}>{fmt(margin)}</td>
                      <td className={`right mono ${margin >= 0 ? "tone-ok-text" : "tone-debt-text"}`}>{marginPct.toFixed(1)}%</td>
                    </tr>
                  );
                })}
                {products.length === 0 && <tr><td colSpan={5} className="empty">Mahsulot yo'q</td></tr>}
              </tbody>
            </table>
          </div>
        );
      })()}

      {report === "pl" && (() => {
        const inv = realInvoices.filter((i) => inRange(i.date));
        const revenue = inv.reduce((s, i) => s + Number(i.total), 0);
        let cogs = 0;
        inv.forEach((i) => {
          (i.items || []).forEach((it) => {
            const p = products.find((x) => x.id === it.product_id);
            cogs += Number(it.qty) * Number(p?.cost_price || 0);
          });
        });
        const netProfit = revenue - cogs;
        const grossMarginPct = revenue > 0 ? (netProfit / revenue) * 100 : 0;
        const otherIncomeTx = cashTransactions.filter((t) => inRange(t.date) && t.type === "kirim" && t.category === "Boshqa tushum");
        const otherExpenseTx = cashTransactions.filter((t) => inRange(t.date) && t.type === "chiqim" && t.category === "Boshqa xarajat");
        const otherIncome = otherIncomeTx.reduce((s, t) => s + Number(t.amount), 0);
        const otherExpense = otherExpenseTx.reduce((s, t) => s + Number(t.amount), 0);
        const finalProfit = netProfit + otherIncome - otherExpense;

        let prevRevenue = null, prevProfit = null, prevRangeLabel = "";
        if (from) {
          const toDate = new Date(to || todayISO());
          const fromDate = new Date(from);
          const rangeDays = Math.max(1, Math.round((toDate - fromDate) / 86400000) + 1);
          const prevTo = new Date(fromDate.getTime() - 86400000);
          const prevFrom = new Date(prevTo.getTime() - (rangeDays - 1) * 86400000);
          const prevFromStr = prevFrom.toISOString().slice(0, 10);
          const prevToStr = prevTo.toISOString().slice(0, 10);
          prevRangeLabel = `${fmtDate(prevFromStr)} — ${fmtDate(prevToStr)}`;
          const prevInv = realInvoices.filter((i) => i.date >= prevFromStr && i.date <= prevToStr);
          prevRevenue = prevInv.reduce((s, i) => s + Number(i.total), 0);
          let prevCogs = 0;
          prevInv.forEach((i) => (i.items || []).forEach((it) => {
            const p = products.find((x) => x.id === it.product_id);
            prevCogs += Number(it.qty) * Number(p?.cost_price || 0);
          }));
          prevProfit = prevRevenue - prevCogs;
        }
        const pctChange = (curr, prev) => (prev === null || prev === 0) ? null : ((curr - prev) / Math.abs(prev)) * 100;
        const revenueChange = pctChange(revenue, prevRevenue);
        const profitChange = pctChange(netProfit, prevProfit);
        const ChangeTag = ({ v }) => v === null ? null : (
          <span className={v >= 0 ? "tone-ok-text" : "tone-debt-text"} style={{ marginLeft: 8, fontSize: 12 }}>
            {v >= 0 ? "▲" : "▼"} {Math.abs(v).toFixed(1)}% oldingi davrga nisbatan
          </span>
        );

        return (
          <div className="card">
            <div className="card-title">Profit & Loss hisoboti</div>
            <div className="muted" style={{ fontSize: 12, marginBottom: 8 }}>
              Tannarx (COGS) tarkibida xom ashyo qiymati va oxirgi 30 kunlik operatsion xarajatlar (ijara, ish haqi va h.k.) ulushi ham bor — shuning uchun bu yerda alohida yana ayirilmaydi.
              {prevRangeLabel && <div style={{ marginTop: 4 }}>Taqqoslanayotgan oldingi davr: {prevRangeLabel}</div>}
            </div>
            <table className="table">
              <tbody>
                <tr><td>Tushum (sotuv)</td><td className="right mono">{fmt(revenue)}<ChangeTag v={revenueChange} /></td></tr>
                <tr><td>Tannarx (COGS, operatsion xarajat bilan)</td><td className="right mono tone-debt-text">−{fmt(cogs)}</td></tr>
                <tr className="editing-row"><td><b>Sof foyda (COGS'dan keyin)</b></td><td className="right mono"><b>{fmt(netProfit)}</b><ChangeTag v={profitChange} /></td></tr>
                <tr><td>Yalpi marja %</td><td className="right mono">{grossMarginPct.toFixed(1)}%</td></tr>
                <tr><td>Boshqa tushumlar</td><td className="right mono tone-ok-text">+{fmt(otherIncome)}</td></tr>
                <tr><td>Boshqa xarajatlar (kategoriyaga kirmagan)</td><td className="right mono tone-debt-text">−{fmt(otherExpense)}</td></tr>
                <tr className="editing-row"><td><b>Yakuniy sof foyda</b></td><td className="right mono"><b className={finalProfit >= 0 ? "tone-ok-text" : "tone-debt-text"}>{fmt(finalProfit)}</b></td></tr>
              </tbody>
            </table>
          </div>
        );
      })()}

      {report === "customers" && (() => {
        const rows = customers
          .filter((c) => !customers.some((x) => x.commission_customer_id === c.id))
          .map((c) => {
            const inv = realInvoices.filter((i) => i.customer_id === c.id && inRange(i.date));
            const revenue = inv.reduce((s, i) => s + Number(i.total), 0);
            let profit = 0;
            inv.forEach((i) => (i.items || []).forEach((it) => {
              const p = products.find((x) => x.id === it.product_id);
              profit += Number(it.qty) * (Number(it.price) - Number(p?.cost_price || 0));
            }));
            const b = balances[c.id] || { invoiced: 0, paid: 0 };
            return { customer: c, revenue, profit, balance: b.invoiced - b.paid };
          })
          .filter((r) => r.revenue > 0)
          .sort((a, b) => b.revenue - a.revenue);
        const totalRevenue = rows.reduce((s, r) => s + r.revenue, 0);
        return (
          <div className="card">
            <div className="card-title">Mijozlar kesimida tahlil</div>
            <table className="table">
              <thead><tr><th>#</th><th>Mijoz</th><th className="right">Tushum</th><th className="right">Ulush</th><th className="right">Foyda</th><th className="right">Marja %</th><th className="right">Joriy qarz</th></tr></thead>
              <tbody>
                {rows.map((r, idx) => {
                  const share = totalRevenue > 0 ? (r.revenue / totalRevenue) * 100 : 0;
                  const marginPct = r.revenue > 0 ? (r.profit / r.revenue) * 100 : 0;
                  return (
                    <tr key={r.customer.id}>
                      <td className="muted">{idx + 1}</td>
                      <td>{r.customer.name} {idx < 3 && <span title="Top 3 mijoz">⭐</span>}</td>
                      <td className="right mono">{fmt(r.revenue)}</td>
                      <td className="right mono muted">{share.toFixed(1)}%</td>
                      <td className={`right mono ${r.profit >= 0 ? "tone-ok-text" : "tone-debt-text"}`}>{fmt(r.profit)}</td>
                      <td className={`right mono ${marginPct >= 0 ? "tone-ok-text" : "tone-debt-text"}`}>{marginPct.toFixed(1)}%</td>
                      <td className="right mono">{fmt(r.balance)}</td>
                    </tr>
                  );
                })}
                {rows.length === 0 && <tr><td colSpan={7} className="empty">Bu davrda sotuv yo'q</td></tr>}
              </tbody>
            </table>
          </div>
        );
      })()}

      {report === "topProducts" && (() => {
        const agg = {};
        realInvoices.filter((i) => inRange(i.date)).forEach((i) => {
          (i.items || []).forEach((it) => {
            if (!agg[it.name]) agg[it.name] = { name: it.name, unit: it.unit, qty: 0, revenue: 0 };
            agg[it.name].qty += Number(it.qty);
            agg[it.name].revenue += Number(it.qty) * Number(it.price);
          });
        });
        const rows = Object.values(agg).sort((a, b) => b.revenue - a.revenue);
        const totalRevenue = rows.reduce((s, r) => s + r.revenue, 0);
        let cumulative = 0;
        const rowsWithAbc = rows.map((r) => {
          cumulative += r.revenue;
          const cumPct = totalRevenue > 0 ? (cumulative / totalRevenue) * 100 : 0;
          const abc = cumPct <= 80 ? "A" : cumPct <= 95 ? "B" : "C";
          return { ...r, share: totalRevenue > 0 ? (r.revenue / totalRevenue) * 100 : 0, cumPct, abc };
        });
        const abcCounts = { A: rowsWithAbc.filter((r) => r.abc === "A").length, B: rowsWithAbc.filter((r) => r.abc === "B").length, C: rowsWithAbc.filter((r) => r.abc === "C").length };
        return (
          <div className="card">
            <div className="card-title">Eng ko'p sotilgan mahsulotlar — ABC (Pareto) tahlili</div>
            <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>
              <b>A</b> — daromadning 80% ini beruvchi mahsulotlar ({abcCounts.A} ta) — asosiy e'tibor shularda bo'lsin.
              <b> B</b> — keyingi 15% ({abcCounts.B} ta). <b>C</b> — qolgan 5% ({abcCounts.C} ta) — ehtimol assortimentni qisqartirish mumkin.
              <div style={{ marginTop: 4 }}><b>Jamlanma %</b> — mahsulotlar tushum bo'yicha eng yuqoridan pastga saralanib, shu qatorgacha bo'lgan barcha mahsulotlarning tushumi jami tushumning necha foizini tashkil qilishi (masalan 3-qatorda 65% bo'lsa, eng ko'p sotilgan 3 mahsulot jami tushumning 65% ini bergan degani).</div>
            </div>
            <table className="table">
              <thead><tr><th>Mahsulot</th><th className="right">Sotilgan miqdor</th><th className="right">Tushum</th><th className="right">Ulush</th><th className="right">Jamlanma %</th><th>Toifa</th></tr></thead>
              <tbody>
                {rowsWithAbc.map((r, idx) => (
                  <tr key={idx}>
                    <td>{r.name}</td>
                    <td className="right mono">{r.qty} {r.unit}</td>
                    <td className="right mono">{fmt(r.revenue)}</td>
                    <td className="right mono muted">{r.share.toFixed(1)}%</td>
                    <td className="right mono muted">{r.cumPct.toFixed(1)}%</td>
                    <td>
                      <span style={{ fontWeight: 700 }} className={r.abc === "A" ? "tone-ok-text" : r.abc === "C" ? "tone-debt-text" : ""}>{r.abc}</span>
                    </td>
                  </tr>
                ))}
                {rows.length === 0 && <tr><td colSpan={6} className="empty">Bu davrda sotuv yo'q</td></tr>}
              </tbody>
            </table>
          </div>
        );
      })()}

      {report === "rfm" && (() => {
        const today = new Date(todayISO());
        const rows = customers
          .filter((c) => !customers.some((x) => x.commission_customer_id === c.id))
          .map((c) => {
            const custInvoices = realInvoices.filter((i) => i.customer_id === c.id).sort((a, b) => new Date(b.date) - new Date(a.date));
            if (custInvoices.length === 0) return null;
            const lastDate = new Date(custInvoices[0].date);
            const recencyDays = Math.max(0, Math.round((today - lastDate) / (1000 * 60 * 60 * 24)));
            const frequency = custInvoices.length;
            const monetary = custInvoices.reduce((s, i) => s + Number(i.total), 0);
            let segment;
            if (recencyDays > 90) segment = "Uxlab qolgan mijoz";
            else if (recencyDays > 60 && frequency >= 3) segment = "Yo'qotilayotgan mijoz";
            else if (recencyDays <= 30 && frequency >= 5) segment = "Sodiq mijoz";
            else if (frequency <= 1) segment = "Yangi / kam faol";
            else segment = "O'rtacha faol";
            return { customer: c, recencyDays, frequency, monetary, segment };
          })
          .filter(Boolean)
          .sort((a, b) => b.monetary - a.monetary);
        const segClass = (seg) => (seg === "Sodiq mijoz" ? "tone-ok-text" : (seg === "Yo'qotilayotgan mijoz" || seg === "Uxlab qolgan mijoz") ? "tone-debt-text" : "muted");
        return (
          <div className="card">
            <div className="card-title">Mijozlar segmentatsiyasi (RFM)</div>
            <div className="muted" style={{ fontSize: 12, marginBottom: 8 }}>Recency (oxirgi xariddan necha kun), Frequency (xaridlar soni), Monetary (jami summa) asosida</div>
            <table className="table">
              <thead><tr><th>Mijoz</th><th className="right">Oxirgi xariddan (kun)</th><th className="right">Xaridlar soni</th><th className="right">Jami summa</th><th>Segment</th></tr></thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.customer.id}>
                    <td>{r.customer.name}</td>
                    <td className="right mono">{r.recencyDays}</td>
                    <td className="right mono">{r.frequency}</td>
                    <td className="right mono">{fmt(r.monetary)}</td>
                    <td className={segClass(r.segment)} style={{ fontWeight: 600 }}>{r.segment}</td>
                  </tr>
                ))}
                {rows.length === 0 && <tr><td colSpan={5} className="empty">Ma'lumot yo'q</td></tr>}
              </tbody>
            </table>
          </div>
        );
      })()}

      {report === "seasonal" && (() => {
        const byMonth = {};
        realInvoices.forEach((i) => {
          const month = i.date.slice(0, 7);
          if (!byMonth[month]) byMonth[month] = { count: 0, total: 0 };
          byMonth[month].count += 1;
          byMonth[month].total += Number(i.total);
        });
        const months = Object.keys(byMonth).sort();
        const maxTotal = Math.max(...months.map((m) => byMonth[m].total), 1);

        // Xom ashyo (masalan yog' turlari) bo'yicha oylik sotuv — norma rasxod orqali hisoblanadi
        const materialsInNorms = rawMaterials.filter((m) => productNorms.some((n) => n.raw_material_id === m.id));
        const materialMonthly = {}; // materialId -> month -> {qty, revenue}
        materialsInNorms.forEach((m) => { materialMonthly[m.id] = {}; });

        realInvoices.forEach((inv) => {
          const month = inv.date.slice(0, 7);
          (inv.items || []).forEach((it) => {
            if (!it.product_id) return;
            const fp = finishedProducts.find((f) => f.linked_product_id === it.product_id);
            if (!fp) return;
            const norms = productNorms.filter((n) => n.finished_product_id === fp.id);
            norms.forEach((n) => {
              if (!materialMonthly[n.raw_material_id]) return;
              if (!materialMonthly[n.raw_material_id][month]) materialMonthly[n.raw_material_id][month] = { qty: 0, revenue: 0 };
              materialMonthly[n.raw_material_id][month].qty += Number(n.qty_per_unit) * Number(it.qty);
              materialMonthly[n.raw_material_id][month].revenue += Number(it.qty) * Number(it.price);
            });
          });
        });

        return (
          <div className="card">
            <div className="card-title">Oylar bo'yicha sotuv dinamikasi</div>
            <table className="table">
              <thead><tr><th>Oy</th><th className="right">Fakturalar soni</th><th className="right">Summa</th><th>Nisbat</th></tr></thead>
              <tbody>
                {months.map((m) => (
                  <tr key={m}>
                    <td className="mono">{m}</td>
                    <td className="right mono">{byMonth[m].count}</td>
                    <td className="right mono">{fmt(byMonth[m].total)}</td>
                    <td style={{ width: 160 }}>
                      <div style={{ background: "var(--accent)", height: 10, borderRadius: 4, width: `${(byMonth[m].total / maxTotal) * 100}%` }} />
                    </td>
                  </tr>
                ))}
                {months.length === 0 && <tr><td colSpan={4} className="empty">Ma'lumot yo'q</td></tr>}
              </tbody>
            </table>

            {materialsInNorms.length > 0 && (
              <>
                <div className="card-title" style={{ fontSize: 14, marginTop: 20 }}>Xom ashyo turlari bo'yicha oylik sotuv (norma rasxod asosida)</div>
                <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>
                  Sotilgan mahsulotlar tarkibida qancha xom ashyo (masalan yog') ketgani, norma rasxod asosida hisoblab chiqilgan. Kg birlik uchun tonna ham ko'rsatiladi.
                </div>
                {materialsInNorms.map((m) => {
                  const rows = months.filter((mo) => materialMonthly[m.id][mo]);
                  if (rows.length === 0) return null;
                  const isKg = (m.unit || "").toLowerCase() === "kg";
                  return (
                    <div key={m.id} style={{ marginBottom: 16 }}>
                      <div style={{ fontWeight: 600, marginBottom: 6 }}>{m.name}</div>
                      <table className="table">
                        <thead><tr><th>Oy</th><th className="right">Miqdor {isKg ? "(tonna)" : `(${m.unit})`}</th><th className="right">Summa</th></tr></thead>
                        <tbody>
                          {rows.map((mo) => {
                            const d = materialMonthly[m.id][mo];
                            return (
                              <tr key={mo}>
                                <td className="mono">{mo}</td>
                                <td className="right mono">{isKg ? (d.qty / 1000).toFixed(2) : fmtQty(d.qty)}</td>
                                <td className="right mono">{fmt(d.revenue)}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  );
                })}
              </>
            )}
          </div>
        );
      })()}

      {report === "agentRating" && (() => {
        const agents = customers.filter((c) => c.is_agent);
        const rows = agents.map((a) => {
          const salesInvoices = realInvoices.filter((i) => i.customer_id === a.id);
          const revenue = salesInvoices.reduce((s, i) => s + Number(i.total), 0);
          const commissionInvoices = invoices.filter((i) => i.customer_id === a.commission_customer_id);
          const commissionEarned = commissionInvoices.reduce((s, i) => s + Number(i.total), 0);
          const commissionPaid = payments.filter((p) => p.customer_id === a.commission_customer_id).reduce((s, p) => s + Number(p.amount), 0);
          const commissionDebt = commissionEarned - commissionPaid;
          return { agent: a, revenue, commissionEarned, commissionDebt };
        }).sort((x, y) => y.revenue - x.revenue);
        return (
          <div className="card">
            <div className="card-title">Agentlar reytingi</div>
            <table className="table">
              <thead><tr><th>Agent</th><th className="right">Sotuv (aylanma)</th><th className="right">Ishlagan komissiya</th><th className="right">To'lanmagan qarz</th></tr></thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.agent.id}>
                    <td>{r.agent.name}</td>
                    <td className="right mono">{fmt(r.revenue)}</td>
                    <td className="right mono">{fmt(r.commissionEarned)}</td>
                    <td className="right mono tone-debt-text">{fmt(r.commissionDebt)}</td>
                  </tr>
                ))}
                {rows.length === 0 && <tr><td colSpan={4} className="empty">Agent sifatida belgilangan mijoz yo'q</td></tr>}
              </tbody>
            </table>
          </div>
        );
      })()}

      {report === "pricingCalc" && (() => {
        const p = products.find((x) => x.id === pcProductId);
        const fp = p ? finishedProducts.find((f) => f.linked_product_id === p.id) : null;
        const norms = fp ? productNorms.filter((n) => n.finished_product_id === fp.id) : [];
        const overhead = computeOverheadPerLiter(cashTransactions, productionBatches, finishedProducts, fixedAssets, 30);
        const productOverhead = overhead.overheadPerLiter * (Number(fp?.volume_liters) || 1);

        const materialRows = norms.map((n) => {
          const m = rawMaterials.find((x) => x.id === n.raw_material_id);
          const defaultCost = avgMaterialCost(n.raw_material_id, rawMaterialBatches);
          const testCost = pcMaterialPrices[n.raw_material_id] !== undefined && pcMaterialPrices[n.raw_material_id] !== ""
            ? Number(pcMaterialPrices[n.raw_material_id]) : defaultCost;
          return { material: m, qtyPerUnit: Number(n.qty_per_unit), defaultCost, testCost, lineCost: Number(n.qty_per_unit) * testCost };
        });

        const materialCost = materialRows.reduce((s, r) => s + r.lineCost, 0);
        const computedCost = materialRows.length > 0 ? materialCost + productOverhead : (p ? Number(p.cost_price) || 0 : 0);
        const price = pcPrice !== "" ? Number(pcPrice) : (p ? Number(p.price) : 0);
        const margin = price - computedCost;
        const marginPct = price > 0 ? (margin / price) * 100 : 0;

        return (
          <div className="card">
            <div className="card-title">Narxlash stsenariysi kalkulyatori</div>
            <div className="form-row wrap" style={{ marginBottom: 12 }}>
              <select className="input" value={pcProductId} onChange={(e) => { setPcProductId(e.target.value); setPcPrice(""); setPcMaterialPrices({}); }}>
                <option value="">— Mahsulotni tanlang —</option>
                {products.map((pr) => <option key={pr.id} value={pr.id}>{pr.name}</option>)}
              </select>
              <input className="input sm mono" type="number" placeholder="Sinov sotish narxi (bo'sh — joriy narx)" value={pcPrice} onChange={(e) => setPcPrice(e.target.value)} />
            </div>

            {p ? (
              <>
                {materialRows.length > 0 && (
                  <>
                    <div className="card-title" style={{ fontSize: 13 }}>Xom ashyo narxlarini test qilish (norma rasxod asosida)</div>
                    <table className="table" style={{ marginBottom: 14 }}>
                      <thead><tr><th>Xom ashyo</th><th className="right">Norma (1 {p.unit} uchun)</th><th className="right">Joriy narx</th><th className="right">Sinov narxi</th><th className="right">Tannarxga ta'siri</th></tr></thead>
                      <tbody>
                        {materialRows.map((r) => (
                          <tr key={r.material?.id}>
                            <td>{r.material?.name}</td>
                            <td className="right mono">{fmtQty(r.qtyPerUnit)} {r.material?.unit}</td>
                            <td className="right mono muted">{fmt(r.defaultCost)}</td>
                            <td className="right">
                              <input
                                className="input mono sm" type="number" placeholder={fmtNum(r.defaultCost)}
                                value={pcMaterialPrices[r.material?.id] ?? ""}
                                onChange={(e) => setPcMaterialPrices({ ...pcMaterialPrices, [r.material.id]: e.target.value })}
                              />
                            </td>
                            <td className="right mono">{fmt(r.lineCost)}</td>
                          </tr>
                        ))}
                        <tr>
                          <td colSpan={4}>Operatsion xarajat ulushi ({fmtQty(fp?.volume_liters || 1)} L × {fmt(overhead.overheadPerLiter)}/L, oxirgi {overhead.days} kun)</td>
                          <td className="right mono">{fmt(productOverhead)}</td>
                        </tr>
                      </tbody>
                    </table>
                    <div className="muted" style={{ fontSize: 12, marginBottom: 14 }}>
                      Sinov narxi bo'sh qoldirilsa, xom ashyoning joriy (faol partiyalar bo'yicha o'rtacha) narxi ishlatiladi. Operatsion xarajat ulushi — oxirgi {overhead.days} kunda "Kassa"da qayd etilgan xarajatlarning shu davrda ishlab chiqarilgan jami litrga bo'lingan qismi, mahsulotning o'z litr hajmiga (1 dona = {fmtQty(fp?.volume_liters || 1)} L) mutanosib hisoblanadi.
                      {overhead.totalLiters === 0 && <span> (hozircha bu davrda ishlab chiqarish bo'lmagani uchun 0 ko'rsatilyapti)</span>}
                    </div>
                  </>
                )}
                {materialRows.length === 0 && (
                  <div className="muted" style={{ fontSize: 12, marginBottom: 12 }}>
                    Bu mahsulot uchun norma rasxod topilmadi — hozirgi saqlangan tannarx (operatsion xarajat allaqachon shu ichida) ishlatilyapti. Xom ashyo asosida test qilish uchun "Ishlab chiqarish → Norma rasxod" bo'limida norma belgilang.
                  </div>
                )}

                <table className="table">
                  <tbody>
                    <tr><td>{materialRows.length > 0 ? "Hisoblangan tannarx (xom ashyo + operatsion xarajat)" : "Tannarx"}</td><td className="right mono">{fmt(computedCost)}</td></tr>
                    <tr><td>Sotish narxi (sinov)</td><td className="right mono">{fmt(price)}</td></tr>
                    <tr><td>Foyda (1 {p.unit} uchun)</td><td className={`right mono ${margin >= 0 ? "tone-ok-text" : "tone-debt-text"}`}>{fmt(margin)}</td></tr>
                    <tr><td>Foyda %</td><td className={`right mono ${margin >= 0 ? "tone-ok-text" : "tone-debt-text"}`}>{marginPct.toFixed(1)}%</td></tr>
                  </tbody>
                </table>
              </>
            ) : <div className="empty">Mahsulotni tanlang</div>}
          </div>
        );
      })()}

      {report === "debtAging" && (() => {
        const today = new Date(todayISO());
        const BUCKETS = [
          { id: "0-30", label: "0–30 kun", min: 0, max: 30 },
          { id: "31-60", label: "31–60 kun", min: 31, max: 60 },
          { id: "61-90", label: "61–90 kun", min: 61, max: 90 },
          { id: "90+", label: "90+ kun", min: 91, max: Infinity },
        ];
        const rows = customers
          .filter((c) => !customers.some((x) => x.commission_customer_id === c.id))
          .map((c) => {
            const custInvoices = [...realInvoices.filter((i) => i.customer_id === c.id)].sort((a, b) => new Date(a.date) - new Date(b.date));
            let remainingPaid = payments.filter((p) => p.customer_id === c.id).reduce((s, p) => s + Number(p.amount), 0);
            let oldestUnpaidDate = null;
            let totalUnpaid = 0;
            for (const inv of custInvoices) {
              let due = Number(inv.total);
              if (remainingPaid > 0) {
                const applied = Math.min(remainingPaid, due);
                due -= applied;
                remainingPaid -= applied;
              }
              if (due > 0.01) {
                totalUnpaid += due;
                if (!oldestUnpaidDate) oldestUnpaidDate = inv.date;
              }
            }
            if (totalUnpaid <= 0.01) return null;
            const days = Math.round((today - new Date(oldestUnpaidDate)) / (1000 * 60 * 60 * 24));
            const bucket = BUCKETS.find((b) => days >= b.min && days <= b.max) || BUCKETS[BUCKETS.length - 1];
            return { customer: c, totalUnpaid, oldestUnpaidDate, days, bucket };
          })
          .filter(Boolean)
          .sort((a, b) => b.days - a.days);

        const bucketTotals = BUCKETS.map((b) => ({ ...b, total: rows.filter((r) => r.bucket.id === b.id).reduce((s, r) => s + r.totalUnpaid, 0) }));

        return (
          <div className="card">
            <div className="card-title">Qarzdorlik muddati (aging)</div>
            <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>
              To'lovlar eng eski fakturadan boshlab hisoblanadi (FIFO); qaysi faktura hali yopilmagan bo'lsa, shu sanadan hozirgacha necha kun o'tgani ko'rsatiladi
            </div>
            <div className="stat-grid" style={{ marginBottom: 16 }}>
              {bucketTotals.map((b) => (
                <StatCard key={b.id} label={b.label} value={fmt(b.total)} mono tone={b.id === "90+" || b.id === "61-90" ? "debt" : undefined} />
              ))}
            </div>
            <table className="table">
              <thead><tr><th>Mijoz</th><th className="right">Qarz summasi</th><th>Eng eski qarz sanasi</th><th className="right">Necha kun</th><th>Toifa</th></tr></thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.customer.id}>
                    <td>{r.customer.name}</td>
                    <td className="right mono">{fmt(r.totalUnpaid)}</td>
                    <td className="muted">{fmtDate(r.oldestUnpaidDate)}</td>
                    <td className="right mono">{r.days}</td>
                    <td className={r.bucket.id === "90+" ? "tone-debt-text" : r.bucket.id === "61-90" ? "tone-debt-text" : "muted"} style={{ fontWeight: 600 }}>{r.bucket.label}</td>
                  </tr>
                ))}
                {rows.length === 0 && <tr><td colSpan={5} className="empty">Qarzdorlik yo'q</td></tr>}
              </tbody>
            </table>
          </div>
        );
      })()}
    </div>
  );
}

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root {
  --bg: #EEF1EC;
  --surface: #FFFFFF;
  --ink: #1F2A24;
  --muted: #6B7570;
  --accent: #1F6F54;
  --accent-dark: #16523E;
  --debt: #B3402F;
  --border: #D9DED6;
}

* { box-sizing: border-box; }
.app-shell { min-height: 100vh; background: var(--bg); font-family: 'Inter', sans-serif; color: var(--ink); }

.topbar { padding: 20px 24px 8px; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; }
.brand { display: flex; align-items: center; gap: 12px; }
.brand-mark { font-family: 'Fraunces', serif; font-size: 28px; font-weight: 700; color: var(--accent); background: var(--surface); border: 1px solid var(--border); width: 44px; height: 44px; display: flex; align-items: center; justify-content: center; border-radius: 10px; }
.brand-title { font-family: 'Fraunces', serif; font-weight: 700; font-size: 19px; }
.brand-sub { font-size: 12px; color: var(--muted); letter-spacing: 0.03em; }

.tabs { display: flex; gap: 4px; padding: 8px 20px; overflow-x: auto; border-bottom: 1px solid var(--border); }
.tab { display: flex; align-items: center; gap: 6px; padding: 9px 14px; border: none; background: transparent; border-radius: 8px 8px 0 0; font-size: 13.5px; font-weight: 500; color: var(--muted); cursor: pointer; white-space: nowrap; }
.tab:hover { background: rgba(31,111,84,0.08); color: var(--ink); }
.tab.active { background: var(--surface); color: var(--accent-dark); box-shadow: 0 -2px 0 var(--accent) inset; font-weight: 600; }
.tab:focus-visible { outline: 2px solid var(--accent); outline-offset: -2px; }

.main { padding: 20px; max-width: 980px; margin: 0 auto; }

.card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 18px; margin-bottom: 16px; }
.card-title { font-family: 'Fraunces', serif; font-weight: 600; font-size: 15px; margin-bottom: 12px; display: flex; align-items: center; gap: 6px; color: var(--accent-dark); }
.card-title-row { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; }

.stat-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 16px; }
@media (max-width: 720px) { .stat-grid { grid-template-columns: repeat(2, 1fr); } }

.balance-columns { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; align-items: start; }
@media (max-width: 640px) { .balance-columns { grid-template-columns: 1fr; } }

.ledger-columns { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; align-items: start; }
@media (max-width: 900px) { .ledger-columns { grid-template-columns: 1fr; } }
.balance-col-title { font-family: 'Fraunces', serif; font-weight: 600; font-size: 12.5px; letter-spacing: 0.06em; color: var(--muted); margin-bottom: 6px; }
.balance-net { display: flex; align-items: center; justify-content: flex-end; gap: 10px; margin-top: 16px; padding-top: 14px; border-top: 1px dashed var(--border); font-size: 15px; }
.balance-net-amount { font-weight: 700; font-size: 19px; }
.stat-card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; }
.stat-label { font-size: 12px; color: var(--muted); margin-bottom: 6px; }
.stat-value { font-family: 'Fraunces', serif; font-size: 21px; font-weight: 700; }
.stat-card.tone-debt .stat-value { color: var(--debt); }
.stat-card.tone-ok .stat-value { color: var(--accent); }

.warn-card { border-color: #E7C9A8; background: #FBF3E8; }
.plain-list { margin: 0; padding-left: 18px; font-size: 13.5px; }

.form-row { display: flex; gap: 8px; align-items: center; }
.form-row.wrap { flex-wrap: wrap; }
.input { font-family: 'Inter', sans-serif; border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px; font-size: 13.5px; background: #fff; color: var(--ink); flex: 1; min-width: 140px; }
.input.sm { flex: 0 0 120px; min-width: 100px; }
.input.xs { flex: 0 0 90px; min-width: 80px; }
.input:focus, select.input:focus { outline: 2px solid var(--accent); outline-offset: 0; border-color: var(--accent); }
.mono { font-family: 'IBM Plex Mono', monospace; }

.btn { display: inline-flex; align-items: center; gap: 6px; border: none; border-radius: 8px; padding: 9px 14px; font-size: 13px; font-weight: 600; cursor: pointer; white-space: nowrap; }
.btn-primary { background: var(--accent); color: #fff; }
.btn-primary:hover { background: var(--accent-dark); }
.btn-primary:disabled { opacity: 0.6; cursor: default; }
.btn-ghost { background: transparent; color: var(--accent-dark); border: 1px solid var(--border); }
.btn-ghost:hover { background: rgba(31,111,84,0.08); }
.btn:focus-visible { outline: 2px solid var(--accent-dark); outline-offset: 2px; }

.table { width: 100%; border-collapse: collapse; font-size: 13.5px; }
.table th { text-align: left; font-weight: 600; color: var(--muted); font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.04em; padding: 8px 8px; border-bottom: 1px solid var(--border); }
.table td { padding: 9px 8px; border-bottom: 1px solid #EEF0EC; vertical-align: middle; }
.table .right { text-align: right; }
.table tr:last-child td { border-bottom: none; }
.row-warn { background: #FDF6ED; }
.row-actions { display: flex; gap: 4px; justify-content: flex-end; }
.icon-btn { border: none; background: transparent; padding: 5px; border-radius: 6px; cursor: pointer; color: var(--muted); display: inline-flex; }
.icon-btn:hover { background: rgba(31,111,84,0.1); color: var(--accent-dark); }
.icon-btn.danger:hover { background: rgba(179,64,47,0.1); color: var(--debt); }
.empty { text-align: center; color: var(--muted); padding: 18px 0; }
.muted { color: var(--muted); }
.editing-row { background: #F2F7F4; }

.search-box { display: flex; align-items: center; gap: 6px; border: 1px solid var(--border); border-radius: 8px; padding: 6px 10px; color: var(--muted); }
.search-box input { border: none; outline: none; font-size: 13px; background: transparent; }

.error-box { display: flex; align-items: center; gap: 6px; background: #FBECE9; color: var(--debt); border: 1px solid #EFCFC7; padding: 8px 12px; border-radius: 8px; font-size: 13px; margin-top: 10px; }

.invoice-footer { display: flex; align-items: center; justify-content: space-between; margin-top: 14px; padding-top: 12px; border-top: 1px dashed var(--border); }
.total-line { font-size: 15px; }
.total-amount { font-weight: 700; font-size: 17px; color: var(--accent-dark); }

.settings-view { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }

.clickable-row { cursor: pointer; }
.clickable-row:hover { background: #F5F8F5; }
.tone-debt-text { color: var(--debt); font-weight: 600; }
.tone-ok-text { color: var(--accent); font-weight: 600; }
.history-cell { background: #FAFBF9; padding: 10px 16px; }
.sub-table td { border-bottom: 1px solid #F0F2ED; font-size: 12.5px; }

.modal-backdrop { position: fixed; inset: 0; background: rgba(20,25,22,0.45); z-index: 55; display: flex; align-items: center; justify-content: center; padding: 20px; }
.modal-card { background: #fff; border-radius: 12px; padding: 20px; width: 100%; max-width: 360px; box-shadow: 0 12px 32px rgba(0,0,0,0.18); }

.print-backdrop { position: fixed; inset: 0; background: rgba(20,25,22,0.55); z-index: 50; overflow-y: auto; padding: 24px; }
.print-toolbar { max-width: 700px; margin: 0 auto 12px; display: flex; justify-content: flex-end; gap: 8px; }
.invoice-sheet { max-width: 700px; margin: 0 auto; background: #fff; padding: 36px; border-radius: 8px; }
.invoice-head { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid var(--ink); padding-bottom: 14px; margin-bottom: 14px; }
.invoice-company { font-family: 'Fraunces', serif; font-weight: 700; font-size: 19px; }
.invoice-meta { text-align: right; font-size: 13px; }
.invoice-to { margin-bottom: 14px; font-size: 13.5px; }
.invoice-table th, .invoice-table td { border-bottom: 1px solid #ddd; }
.invoice-total { text-align: right; font-size: 16px; font-weight: 700; margin-top: 12px; }
.invoice-sign-row { display: flex; justify-content: space-between; margin-top: 40px; font-size: 13px; }

@media print {
  body * { visibility: hidden; }
  #print-area, #print-area * { visibility: visible; }
  #print-area { position: absolute; top: 0; left: 0; width: 100%; padding: 0; }
  .no-print { display: none !important; }
}
`;
