import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell,
} from "recharts";
import {
  LayoutGrid, BarChart3, FileText, Users, Settings, Search, Bell, Moon,
  Download, ChevronDown, Radio, Pause, Play, TrendingUp, TrendingDown,
  PackageCheck, Clock3, ShoppingBag, Truck, CheckCircle2, XCircle, Megaphone,
  AlertTriangle,
} from "lucide-react";

// ====================================================================
// KONFIGURASI — diambil dari environment variable Vercel/Vite.
// Set VITE_SUPABASE_URL dan VITE_SUPABASE_ANON_KEY di Vercel Project Settings → Environment Variables
// ====================================================================
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "https://YOUR-PROJECT-REF.supabase.co";
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || "YOUR-ANON-PUBLIC-KEY";

const IS_CONFIGURED =
  !SUPABASE_URL.includes("YOUR-PROJECT-REF") && !SUPABASE_ANON_KEY.includes("YOUR-ANON-PUBLIC-KEY");

const REST = `${SUPABASE_URL}/rest/v1`;
const REST_HEADERS = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
};

async function pgFetch(path) {
  const res = await fetch(`${REST}/${path}`, { headers: REST_HEADERS });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Supabase fetch gagal (${res.status}): ${text || res.statusText}`);
  }
  return res.json();
}

// ---------- helpers ----------
const rupiah = (n) => "Rp" + Math.round(n || 0).toLocaleString("id-ID");
const compact = (n) =>
  n >= 1_000_000 ? (n / 1_000_000).toFixed(1) + "jt" :
  n >= 1_000 ? (n / 1_000).toFixed(1) + "rb" : String(Math.round(n || 0));

function dayKey(iso) {
  return new Date(iso).toISOString().slice(0, 10); // YYYY-MM-DD
}

// Traffic & Ads Spend belum ada sumber data resmi dari Shopee Open API standar
// (butuh scope Business Insight / Shopee Ads API terpisah). Ditandai "estimasi".
const TRAFFIC_SOURCES = [
  { name: "Shopee Ads", value: 34, color: "#8b5cf6" },
  { name: "Pencarian Organik", value: 26, color: "#22d3ee" },
  { name: "Shopee Live", value: 18, color: "#34d399" },
  { name: "Affiliate / Konten Kreator", value: 13, color: "#fbbf24" },
  { name: "Direct / Follower Toko", value: 9, color: "#f472b6" },
];
const ESTIMATED_TRAFFIC_PER_ORDER = 35; // asumsi kasar, ganti kalau sudah ada data traffic asli
const ADS_SPEND_PCT_OF_GMV = 0.08; // asumsi kasar, ganti kalau sudah integrasi Shopee Ads API

const RANGE_OPTIONS = [
  { key: "today", label: "Hari Ini", days: 1 },
  { key: "7d", label: "7 Hari Terakhir", days: 7 },
  { key: "30d", label: "30 Hari Terakhir", days: 30 },
  { key: "month", label: "Bulan Ini", days: 30 },
];

const STATUS_META = {
  READY_TO_SHIP: { label: "Perlu Dikirim", icon: PackageCheck, color: "#f59e0b" },
  SHIPPED: { label: "Sedang Dikirim", icon: Truck, color: "#6366f1" },
  TO_CONFIRM_RECEIVE: { label: "Sedang Dikirim", icon: Truck, color: "#6366f1" },
  COMPLETED: { label: "Selesai", icon: CheckCircle2, color: "#22c55e" },
  CANCELLED: { label: "Batal / Retur", icon: XCircle, color: "#ef4444" },
  IN_CANCEL: { label: "Batal / Retur", icon: XCircle, color: "#ef4444" },
  TO_RETURN: { label: "Batal / Retur", icon: XCircle, color: "#ef4444" },
  UNPAID: { label: "Belum Dibayar", icon: Clock3, color: "#94a3b8" },
};
function statusMeta(status) {
  return STATUS_META[status] || { label: status || "Lainnya", icon: PackageCheck, color: "#64748b" };
}

// ---------- small components ----------
function Sparkline({ data, color }) {
  const w = 90, h = 30;
  if (!data || data.length < 2) return <svg width={w} height={h} />;
  const max = Math.max(...data), min = Math.min(...data);
  const pts = data.map((v, i) => {
    const x = (i / (data.length - 1)) * w;
    const y = h - ((v - min) / (max - min || 1)) * h;
    return `${x},${y}`;
  }).join(" ");
  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function StatCard({ label, value, delta, positiveIsGood, sparkData, color, icon: Icon, estimated }) {
  const hasDelta = typeof delta === "number";
  const isUp = hasDelta && delta >= 0;
  const good = positiveIsGood ? isUp : !isUp;
  return (
    <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col gap-4">
      <div className="flex items-start justify-between">
        <div className="flex items-center gap-2 text-slate-400 text-xs font-medium tracking-wide uppercase">
          <Icon size={14} />
          {label}
          {estimated && <span className="text-[9px] normal-case text-slate-600">(estimasi)</span>}
        </div>
        <Sparkline data={sparkData} color={color} />
      </div>
      <div className="text-2xl font-semibold text-slate-50">{value}</div>
      {hasDelta ? (
        <div className={"inline-flex w-fit items-center gap-1 rounded-full px-2 py-1 text-xs font-medium " +
          (good ? "bg-emerald-950 text-emerald-400" : "bg-rose-950 text-rose-400")}>
          {isUp ? <TrendingUp size={12} /> : <TrendingDown size={12} />}
          {Math.abs(delta).toFixed(1)}% vs periode sebelumnya
        </div>
      ) : (
        <div className="text-xs text-slate-600">&nbsp;</div>
      )}
    </div>
  );
}

function SidebarItem({ icon: Icon, label, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className={"w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors " +
        (active ? "bg-violet-600/15 text-violet-300" : "text-slate-400 hover:text-slate-200 hover:bg-slate-900")}
    >
      <Icon size={18} />
      {label}
    </button>
  );
}

export default function SellerDashboard() {
  const [now, setNow] = useState(new Date());
  const [isLive, setIsLive] = useState(true);
  const [range, setRange] = useState(RANGE_OPTIONS[2]);
  const [rangeOpen, setRangeOpen] = useState(false);
  const [activeMenu, setActiveMenu] = useState("Overview");

  const [orders, setOrders] = useState([]);
  const [orderItems, setOrderItems] = useState([]);
  const [liveFeed, setLiveFeed] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const pollRef = useRef(null);

  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  const sinceISO = useMemo(() => {
    const days = range.key === "today" ? 1 : range.days;
    const d = new Date();
    d.setDate(d.getDate() - days);
    return d.toISOString();
  }, [range]);

  const loadData = useCallback(async () => {
    if (!IS_CONFIGURED) {
      setLoading(false);
      return;
    }
    try {
      setError(null);
      const [ordersRes, itemsRes, feedRes] = await Promise.all([
        pgFetch(`orders?select=order_sn,status,total_amount,create_time,buyer_username&create_time=gte.${sinceISO}&order=create_time.desc&limit=2000`),
        pgFetch(`order_items?select=order_sn,item_name,sku,qty,subtotal,orders(create_time,status)&limit=3000&order=order_sn.desc`),
        pgFetch(`orders?select=order_sn,buyer_username,total_amount,create_time,order_items(item_name,qty)&status=eq.READY_TO_SHIP&order=create_time.desc&limit=8`),
      ]);
      setOrders(ordersRes || []);
      setOrderItems(itemsRes || []);
      setLiveFeed(feedRes || []);
    } catch (e) {
      setError(e.message || String(e));
    } finally {
      setLoading(false);
    }
  }, [sinceISO]);

  useEffect(() => {
    loadData();
  }, [loadData]);

  useEffect(() => {
    if (!isLive || !IS_CONFIGURED) {
      if (pollRef.current) clearInterval(pollRef.current);
      return;
    }
    pollRef.current = setInterval(loadData, 8000); // polling tiap 8 detik selagi Live aktif
    return () => clearInterval(pollRef.current);
  }, [isLive, loadData]);

  // ---------- derived data ----------
  const totalGMV = useMemo(() => orders.reduce((s, o) => s + (Number(o.total_amount) || 0), 0), [orders]);
  const totalOrders = orders.length;
  const traffic = totalOrders * ESTIMATED_TRAFFIC_PER_ORDER;
  const convRate = traffic > 0 ? (totalOrders / traffic) * 100 : 0;
  const adsSpend = totalGMV * ADS_SPEND_PCT_OF_GMV;

  const dailyGmvSeries = useMemo(() => {
    const map = {};
    for (const o of orders) {
      const key = dayKey(o.create_time);
      map[key] = (map[key] || 0) + (Number(o.total_amount) || 0);
    }
    return Object.entries(map)
      .sort(([a], [b]) => (a < b ? -1 : 1))
      .map(([day, gmv]) => ({ day: day.slice(5), gmv }));
  }, [orders]);

  const sparkGMV = dailyGmvSeries.length ? dailyGmvSeries.map((d) => d.gmv) : [0, 0];

  const bestSellers = useMemo(() => {
    const map = {};
    for (const it of orderItems) {
      const inRange = it.orders && it.orders.create_time >= sinceISO;
      if (!inRange) continue;
      const key = it.sku || it.item_name;
      if (!map[key]) map[key] = { sku: it.sku || "-", name: it.item_name, qty: 0, revenue: 0 };
      map[key].qty += Number(it.qty) || 0;
      map[key].revenue += Number(it.subtotal) || 0;
    }
    return Object.values(map).sort((a, b) => b.revenue - a.revenue).slice(0, 6);
  }, [orderItems, sinceISO]);

  const statusBreakdown = useMemo(() => {
    const map = {};
    for (const o of orders) {
      const key = o.status || "UNKNOWN";
      map[key] = (map[key] || 0) + 1;
    }
    return Object.entries(map)
      .map(([status, count]) => ({ status, count, ...statusMeta(status) }))
      .sort((a, b) => b.count - a.count);
  }, [orders]);
  const totalStatus = statusBreakdown.reduce((s, x) => s + x.count, 0);
  const totalTrafficPct = TRAFFIC_SOURCES.reduce((s, x) => s + x.value, 0);

  function handleExport() {
    const wb = XLSX.utils.book_new();

    const summary = [
      ["Ringkasan Toko", range.label],
      [],
      ["Metrik", "Nilai"],
      ["GMV (Omset)", totalGMV],
      ["Total Order", totalOrders],
      ["Traffic (estimasi)", traffic],
      ["Conversion Rate % (estimasi)", convRate.toFixed(2)],
      ["Ads Spend (estimasi)", adsSpend],
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, ws1, "Ringkasan");

    const gmvRows = [["Tanggal", "GMV"], ...dailyGmvSeries.map((d) => [d.day, d.gmv])];
    const ws2 = XLSX.utils.aoa_to_sheet(gmvRows);
    XLSX.utils.book_append_sheet(wb, ws2, "GMV Harian");

    const bsRows = [["SKU", "Produk", "Qty Terjual", "Revenue (GMV)"],
      ...bestSellers.map((b) => [b.sku, b.name, b.qty, b.revenue])];
    const ws3 = XLSX.utils.aoa_to_sheet(bsRows);
    XLSX.utils.book_append_sheet(wb, ws3, "Produk Terlaris");

    XLSX.writeFile(wb, `dashboard-toko-${range.key}.xlsx`);
  }

  if (!IS_CONFIGURED) {
    return (
      <div className="flex items-center justify-center min-h-screen bg-slate-950 text-slate-300 p-8">
        <div className="max-w-md bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col gap-3">
          <div className="flex items-center gap-2 text-amber-400 font-semibold">
            <AlertTriangle size={18} />
            Belum dikonfigurasi
          </div>
          <p className="text-sm text-slate-400">
            Buka file ini, isi <code className="text-violet-300">SUPABASE_URL</code> dan{" "}
            <code className="text-violet-300">SUPABASE_ANON_KEY</code> di bagian atas kode
            (ambil dari Supabase → Project Settings → API), lalu simpan ulang.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-full min-h-screen bg-slate-950 text-slate-200" style={{ fontFamily: "Inter, system-ui, sans-serif" }}>
      {/* Sidebar */}
      <aside className="w-60 shrink-0 border-r border-slate-800 flex flex-col justify-between py-5 px-4">
        <div>
          <div className="flex items-center gap-2 px-2 mb-8">
            <div className="w-8 h-8 rounded-lg bg-violet-600 flex items-center justify-center">
              <ShoppingBag size={16} className="text-white" />
            </div>
            <span className="font-semibold text-slate-50 text-lg">TokoKu</span>
          </div>
          <div className="text-xs font-semibold text-slate-600 uppercase tracking-wide px-3 mb-2">Workspace</div>
          <nav className="flex flex-col gap-1">
            <SidebarItem icon={LayoutGrid} label="Overview" active={activeMenu === "Overview"} onClick={() => setActiveMenu("Overview")} />
            <SidebarItem icon={BarChart3} label="Menu 2" active={activeMenu === "Menu 2"} onClick={() => setActiveMenu("Menu 2")} />
            <SidebarItem icon={FileText} label="Menu 3" active={activeMenu === "Menu 3"} onClick={() => setActiveMenu("Menu 3")} />
            <SidebarItem icon={Users} label="Menu 4" active={activeMenu === "Menu 4"} onClick={() => setActiveMenu("Menu 4")} />
            <SidebarItem icon={Settings} label="Menu 5" active={activeMenu === "Menu 5"} onClick={() => setActiveMenu("Menu 5")} />
          </nav>
        </div>
        <div className="flex items-center gap-3 border-t border-slate-800 pt-4 px-1">
          <div className="w-8 h-8 rounded-full bg-violet-700 flex items-center justify-center text-xs font-semibold text-white">AR</div>
          <div className="leading-tight">
            <div className="text-sm text-slate-200 font-medium">Toko Anda</div>
            <div className="text-xs text-slate-500">Seller</div>
          </div>
        </div>
      </aside>

      {/* Main */}
      <div className="flex-1 flex flex-col min-w-0">
        {/* Top bar */}
        <header className="flex items-center justify-between px-8 py-4 border-b border-slate-800">
          <div className="text-sm text-slate-500">Toko Anda <span className="mx-1">/</span> <span className="text-slate-200">{activeMenu}</span></div>
          <div className="flex items-center gap-4">
            <div className="hidden md:flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-sm text-slate-500 w-64">
              <Search size={14} />
              Cari pesanan, produk...
            </div>
            <div className="text-sm text-slate-400 tabular-nums w-16 text-right">
              {now.toLocaleTimeString("id-ID", { hour12: false })}
            </div>
            <Bell size={18} className="text-slate-500" />
            <Moon size={18} className="text-slate-500" />
            <div className="w-8 h-8 rounded-full bg-violet-700 flex items-center justify-center text-xs font-semibold text-white">AR</div>
          </div>
        </header>

        <main className="flex-1 overflow-y-auto px-8 py-6 flex flex-col gap-6">
          {error && (
            <div className="bg-rose-950 border border-rose-900 text-rose-300 text-sm rounded-xl px-4 py-3 flex items-center gap-2">
              <AlertTriangle size={14} />
              {error} — cek RLS policy tabel Supabase kamu (anon role perlu izin SELECT).
            </div>
          )}

          {/* Title row */}
          <div className="flex flex-wrap items-start justify-between gap-4">
            <div>
              <h1 className="text-2xl font-semibold text-slate-50">Overview</h1>
              <p className="text-slate-500 text-sm mt-1">
                {loading ? "Memuat data..." : "Selamat datang kembali — data langsung dari Supabase."}
              </p>
            </div>
            <div className="flex items-center gap-3">
              <div className="relative">
                <button
                  onClick={() => setRangeOpen((o) => !o)}
                  className="flex items-center gap-2 bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-sm text-slate-200"
                >
                  <Clock3 size={14} className="text-slate-500" />
                  {range.label}
                  <ChevronDown size={14} className="text-slate-500" />
                </button>
                {rangeOpen && (
                  <div className="absolute right-0 mt-1 w-48 bg-slate-900 border border-slate-800 rounded-lg overflow-hidden z-10 shadow-lg">
                    {RANGE_OPTIONS.map((opt) => (
                      <button
                        key={opt.key}
                        onClick={() => { setRange(opt); setRangeOpen(false); }}
                        className={"w-full text-left px-3 py-2 text-sm hover:bg-slate-800 " +
                          (opt.key === range.key ? "text-violet-300" : "text-slate-300")}
                      >
                        {opt.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>

              <button
                onClick={() => setIsLive((v) => !v)}
                className={"flex items-center gap-2 rounded-lg px-3 py-2 text-sm font-medium border " +
                  (isLive ? "bg-emerald-950 border-emerald-900 text-emerald-400" : "bg-slate-900 border-slate-800 text-slate-400")}
              >
                {isLive ? <Radio size={14} className="animate-pulse" /> : <Pause size={14} />}
                {isLive ? "Live" : "Paused"}
              </button>

              <button
                onClick={handleExport}
                className="flex items-center gap-2 bg-violet-600 hover:bg-violet-500 transition-colors rounded-lg px-3 py-2 text-sm font-medium text-white"
              >
                <Download size={14} />
                Export Excel
              </button>
            </div>
          </div>

          {/* Stat cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
            <StatCard label="GMV (Omset)" value={rupiah(totalGMV)} sparkData={sparkGMV} color="#34d399" icon={ShoppingBag} />
            <StatCard label="Total Order" value={totalOrders.toLocaleString("id-ID")} sparkData={sparkGMV.map((v) => v ? 1 : 0)} color="#818cf8" icon={PackageCheck} />
            <StatCard label="Conversion Rate" value={convRate.toFixed(2) + "%"} sparkData={[convRate, convRate]} color="#f87171" icon={TrendingUp} estimated />
            <StatCard label="Ads Spend" value={rupiah(adsSpend)} sparkData={sparkGMV.map((v) => v * ADS_SPEND_PCT_OF_GMV)} color="#fbbf24" icon={Megaphone} estimated />
          </div>

          {/* Revenue + Live activity */}
          <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
            <div className="xl:col-span-2 bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col">
              <div className="flex items-start justify-between mb-4">
                <div>
                  <h2 className="text-slate-50 font-semibold">GMV Harian</h2>
                  <p className="text-slate-500 text-xs mt-0.5">{range.label} · dari tabel orders</p>
                </div>
                <div className="text-xs text-slate-500">{dailyGmvSeries.length} hari</div>
              </div>
              <div style={{ height: 280 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <AreaChart data={dailyGmvSeries} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                    <defs>
                      <linearGradient id="gmvFill" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#8b5cf6" stopOpacity={0.5} />
                        <stop offset="100%" stopColor="#8b5cf6" stopOpacity={0} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid stroke="#1e293b" vertical={false} />
                    <XAxis dataKey="day" stroke="#475569" fontSize={11} tickLine={false} axisLine={false} />
                    <YAxis stroke="#475569" fontSize={11} tickLine={false} axisLine={false} tickFormatter={compact} width={44} />
                    <Tooltip
                      contentStyle={{ background: "#0f172a", border: "1px solid #1e293b", borderRadius: 8, fontSize: 12 }}
                      labelStyle={{ color: "#94a3b8" }}
                      formatter={(v) => [rupiah(v), "GMV"]}
                    />
                    <Area type="monotone" dataKey="gmv" stroke="#8b5cf6" strokeWidth={2} fill="url(#gmvFill)" />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
              <div className="flex gap-8 mt-4 pt-4 border-t border-slate-800 text-sm">
                <div>
                  <div className="text-slate-500 text-xs">Total Order ({range.label})</div>
                  <div className="text-slate-100 font-medium">{totalOrders}</div>
                </div>
                <div>
                  <div className="text-slate-500 text-xs">Traffic (estimasi)</div>
                  <div className="text-slate-100 font-medium">{traffic.toLocaleString("id-ID")}</div>
                </div>
                <div>
                  <div className="text-slate-500 text-xs">Rata-rata Order Value</div>
                  <div className="text-slate-100 font-medium">{rupiah(totalGMV / Math.max(totalOrders, 1))}</div>
                </div>
              </div>
            </div>

            {/* Live activity - incoming orders ready to ship */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 flex flex-col">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-slate-50 font-semibold flex items-center gap-2">
                  Pesanan Masuk
                  <span className={"w-2 h-2 rounded-full " + (isLive ? "bg-emerald-400 animate-pulse" : "bg-slate-600")} />
                </h2>
                <span className="text-xs text-slate-500">Siap Dikirim</span>
              </div>
              <div className="flex flex-col gap-3 overflow-y-auto pr-1" style={{ maxHeight: 420 }}>
                {liveFeed.length === 0 && !loading && (
                  <div className="text-sm text-slate-600 text-center py-8">Belum ada pesanan siap dikirim.</div>
                )}
                {liveFeed.map((o) => (
                  <div key={o.order_sn} className="flex items-start gap-3 border-b border-slate-800/60 pb-3 last:border-0">
                    <div className="w-8 h-8 rounded-full bg-violet-950 text-violet-300 flex items-center justify-center text-xs font-semibold shrink-0">
                      {(o.buyer_username || "?").slice(0, 1).toUpperCase()}
                    </div>
                    <div className="min-w-0 flex-1">
                      <div className="text-sm text-slate-200 font-medium truncate">{o.order_sn}</div>
                      <div className="text-xs text-slate-500 truncate">
                        {(o.order_items || []).map((it) => `${it.item_name} x${it.qty}`).join(", ") || "-"}
                      </div>
                      <div className="text-[11px] text-slate-600 mt-0.5">
                        {new Date(o.create_time).toLocaleTimeString("id-ID", { hour12: false })} · {o.buyer_username}
                      </div>
                    </div>
                    <div className="text-xs font-medium text-emerald-400 shrink-0">{rupiah(o.total_amount)}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Bottom row */}
          <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
            {/* Traffic sources (estimasi, belum ada sumber data resmi) */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6">
              <h2 className="text-slate-50 font-semibold">Sumber Traffic / Penjualan</h2>
              <p className="text-slate-500 text-xs mt-0.5 mb-4">Estimasi — belum terhubung ke sumber data traffic asli</p>
              <div className="relative flex items-center justify-center" style={{ height: 200 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={TRAFFIC_SOURCES} dataKey="value" nameKey="name" innerRadius={55} outerRadius={80} paddingAngle={3}>
                      {TRAFFIC_SOURCES.map((s, i) => <Cell key={i} fill={s.color} stroke="none" />)}
                    </Pie>
                  </PieChart>
                </ResponsiveContainer>
                <div className="absolute flex flex-col items-center pointer-events-none">
                  <div className="text-xl font-semibold text-slate-50">{traffic.toLocaleString("id-ID")}</div>
                  <div className="text-xs text-slate-500">estimasi kunjungan</div>
                </div>
              </div>
              <div className="flex flex-col gap-2 mt-2">
                {TRAFFIC_SOURCES.map((s) => (
                  <div key={s.name} className="flex items-center justify-between text-xs">
                    <div className="flex items-center gap-2 text-slate-400">
                      <span className="w-2 h-2 rounded-full" style={{ background: s.color }} />
                      {s.name}
                    </div>
                    <span className="text-slate-300 font-medium">{((s.value / totalTrafficPct) * 100).toFixed(0)}%</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Order status — real, dari tabel orders */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6">
              <div className="flex items-center justify-between mb-1">
                <h2 className="text-slate-50 font-semibold">Status Pesanan</h2>
                <span className={"w-2 h-2 rounded-full " + (isLive ? "bg-emerald-400 animate-pulse" : "bg-slate-600")} />
              </div>
              <div className="text-3xl font-semibold text-slate-50 mt-2">{totalStatus.toLocaleString("id-ID")}</div>
              <div className="text-xs text-slate-500 mb-5">total pesanan {range.label.toLowerCase()}</div>
              <div className="flex flex-col gap-4">
                {statusBreakdown.length === 0 && !loading && (
                  <div className="text-sm text-slate-600">Belum ada data pesanan di periode ini.</div>
                )}
                {statusBreakdown.map((s) => {
                  const pct = totalStatus ? (s.count / totalStatus) * 100 : 0;
                  const Icon = s.icon;
                  return (
                    <div key={s.status}>
                      <div className="flex items-center justify-between text-xs mb-1">
                        <div className="flex items-center gap-1.5 text-slate-400">
                          <Icon size={12} style={{ color: s.color }} />
                          {s.label}
                        </div>
                        <span className="text-slate-300 font-medium">{s.count}</span>
                      </div>
                      <div className="h-1.5 bg-slate-800 rounded-full overflow-hidden">
                        <div className="h-full rounded-full" style={{ width: `${pct}%`, background: s.color }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Best sellers — real, agregasi dari order_items */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-slate-50 font-semibold">Produk Terlaris</h2>
                <span className="text-xs text-slate-600">{range.label}</span>
              </div>
              {bestSellers.length === 0 && !loading ? (
                <div className="text-sm text-slate-600 py-8 text-center">Belum ada data penjualan produk.</div>
              ) : (
                <table className="w-full text-xs">
                  <thead>
                    <tr className="text-slate-500 text-left">
                      <th className="font-normal pb-2">Produk</th>
                      <th className="font-normal pb-2 text-right">Qty</th>
                      <th className="font-normal pb-2 text-right">GMV</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bestSellers.map((b) => (
                      <tr key={b.sku + b.name} className="border-t border-slate-800/70">
                        <td className="py-2 pr-2">
                          <div className="text-slate-200 truncate max-w-[140px]">{b.name}</div>
                          <div className="text-slate-600 text-[10px]">{b.sku}</div>
                        </td>
                        <td className="py-2 text-right text-slate-300">{b.qty}</td>
                        <td className="py-2 text-right text-slate-300">{compact(b.revenue)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}
