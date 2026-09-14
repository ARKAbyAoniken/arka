import React, { useMemo, useState, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart3, Package, ClipboardList, Utensils,
  Trash2, WalletCards, FileSpreadsheet, Users, Settings, Search, Plus, Save, Upload, X, Download,
} from "lucide-react";

const modules = [
  { key: "dashboard", label: "Dashboard", icon: BarChart3 },
  { key: "stock", label: "Stocktake", icon: Package },
  { key: "recipes", label: "Recipes", icon: ClipboardList },
  { key: "production", label: "Production", icon: Utensils },
  { key: "waste", label: "Waste", icon: Trash2 },
  { key: "sales", label: "Sales", icon: WalletCards },
  { key: "reports", label: "Reports", icon: FileSpreadsheet },
  { key: "users", label: "Users", icon: Users },
  { key: "settings", label: "Settings", icon: Settings },
];

type Product = { id: number; name: string; category: string; unit: string; price: number; realStock: number; theoreticalStock: number; parLevel: number };
type Ingredient = { id: number; productId: number; quantity: number; unit: string };
type Recipe = { id: number; name: string; category: string; portions: number; prepTime: number; allergens: string[]; procedure: string[]; photo: string | null; ingredients: Ingredient[] };
type ProductionEntry = { id: number; date: string; recipeId: number; recipeName: string; portions: number; notes: string };
type WasteEntry = { id: number; date: string; type: "product" | "production"; refId: number; refName: string; quantity: number; unit: string; reason: string; estimatedCost: number; notes: string };
type SaleItem = { id: number; name: string; category: string; quantity: number; unitPrice: number; total: number };
type SaleEntry = { id: number; date: string; source: string; items: SaleItem[]; totalRevenue: number; notes: string };
type AiProduct = { name: string; quantity: number; unit: string; price: number };

const initialProducts: Product[] = [
  { id: 1, name: "Rioplatense Arg Ribeye", category: "Meat", unit: "kg", price: 29.95, realStock: 134.55, theoreticalStock: 134.55, parLevel: 40 },
  { id: 2, name: "Rioplatense Arg Striploin", category: "Meat", unit: "kg", price: 17.5, realStock: 162.3, theoreticalStock: 162.3, parLevel: 40 },
  { id: 3, name: "Arg Beef Fillet Santa Giulia", category: "Meat", unit: "kg", price: 33.5, realStock: 132, theoreticalStock: 132, parLevel: 40 },
  { id: 4, name: "Flap Meat", category: "Meat", unit: "kg", price: 17.5, realStock: 91.84, theoreticalStock: 91.84, parLevel: 30 },
  { id: 5, name: "Entraña (Inside Skirt)", category: "Meat", unit: "kg", price: 14, realStock: 14.9, theoreticalStock: 14.9, parLevel: 40 },
  { id: 6, name: "Provolone Cheese", category: "Dairy", unit: "kg", price: 11.55, realStock: 20, theoreticalStock: 20, parLevel: 5 },
  { id: 7, name: "Eggs Medium 1x180", category: "Dairy", unit: "units", price: 0.16, realStock: 150, theoreticalStock: 150, parLevel: 180 },
  { id: 8, name: "Chorizo Criollo La Ribera", category: "Meat", unit: "kg", price: 0.07, realStock: 845, theoreticalStock: 845, parLevel: 10 },
  { id: 9, name: "Morcilla", category: "Meat", unit: "each", price: 0, realStock: 247, theoreticalStock: 247, parLevel: 100 },
  { id: 10, name: "Sweet Potato Fries 4x2.27kg", category: "Frozen", unit: "packs", price: 8.99, realStock: 18, theoreticalStock: 18, parLevel: 10 },
];

const initialRecipes: Recipe[] = [
  { id: 1, name: "Empanadas de Carne", category: "Starters", portions: 12, prepTime: 45, allergens: ["Gluten", "Eggs"], procedure: ["Preparar el relleno", "Armar las empanadas", "Hornear a 200°C por 20 minutos"], photo: null, ingredients: [{ id: 1, productId: 1, quantity: 0.5, unit: "kg" }, { id: 2, productId: 7, quantity: 2, unit: "unit" }] },
  { id: 2, name: "Ribeye 300g", category: "Mains", portions: 1, prepTime: 15, allergens: [], procedure: ["Sellar a fuego alto 2 min por lado", "Reposar 5 minutos", "Servir con chimichurri"], photo: null, ingredients: [{ id: 3, productId: 1, quantity: 0.3, unit: "kg" }] },
];

const ALLERGENS = ["Gluten", "Dairy", "Eggs", "Nuts", "Soy", "Fish", "Shellfish", "Sesame"];
const WASTE_REASONS = ["Spoilage", "Over-production", "Damaged", "Expired", "Cooking error", "Dropped", "Other"];
const CATEGORIES = ["Meat", "Dairy", "Frozen", "Fresh Veg", "Dry Goods", "Drinks", "Cleaning", "Other", "Imported"];
const UNITS = ["kg", "g", "l", "ml", "unit", "each", "packs", "litres", "kilos", "bunches"];
const SALE_CATEGORIES = ["Starters", "Mains", "Desserts", "Drinks", "Sides", "Other"];

function money(value: number) {
  return new Intl.NumberFormat("en-GB", { style: "currency", currency: "GBP" }).format(value || 0);
}

function ArkaLogo() {
  return (
    <div className="flex items-center gap-3">
      <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-white/15 text-[#EAD5B3]">
        <svg viewBox="0 0 220 220" className="h-8 w-8">
          <path d="M110 28 L194 178 H26 Z" fill="none" stroke="currentColor" strokeWidth="14" strokeLinejoin="miter" />
          <path d="M72 96 C72 142 88 178 110 178 C132 178 148 142 148 96" fill="none" stroke="currentColor" strokeWidth="14" strokeLinecap="round" />
        </svg>
      </div>
      <div>
        <div className="text-2xl font-black tracking-[0.28em] text-white">ARKA</div>
        <div className="text-xs font-semibold uppercase tracking-[0.18em] text-[#EAD5B3]">Kitchen OS</div>
      </div>
    </div>
  );
}

function Shell({ active, setActive, children }: { active: string; setActive: (k: string) => void; children: React.ReactNode }) {
  const [menuOpen, setMenuOpen] = useState(false);

  return (
    <div className="min-h-screen bg-[#F7F8FA] text-[#2D3748]">
      <div className="flex min-h-screen">
        <aside className="hidden w-72 bg-[#1a3a5c] p-5 lg:block sticky top-0 h-screen overflow-y-auto">
          <ArkaLogo />
          <nav className="mt-10 grid gap-2">
            {modules.map((item) => {
              const Icon = item.icon;
              const selected = active === item.key;
              return (
                <button key={item.key} onClick={() => setActive(item.key)}
                  className={`flex items-center gap-3 rounded-xl px-4 py-3 text-left text-sm font-semibold transition ${selected ? "bg-white text-[#1a3a5c]" : "text-white/75 hover:bg-white/10 hover:text-white"}`}>
                  <Icon size={18} />{item.label}
                </button>
              );
            })}
          </nav>
        </aside>

        {menuOpen && (
          <div className="fixed inset-0 z-50 lg:hidden">
            <div className="absolute inset-0 bg-black/50" onClick={() => setMenuOpen(false)} />
            <aside className="absolute left-0 top-0 h-full w-72 bg-[#1a3a5c] p-5 overflow-y-auto">
              <div className="flex items-center justify-between mb-6">
                <ArkaLogo />
                <button onClick={() => setMenuOpen(false)} className="text-white/75 hover:text-white">
                  <X size={24} />
                </button>
              </div>
              <nav className="grid gap-2">
                {modules.map((item) => {
                  const Icon = item.icon;
                  const selected = active === item.key;
                  return (
                    <button key={item.key} onClick={() => { setActive(item.key); setMenuOpen(false); }}
                      className={`flex items-center gap-3 rounded-xl px-4 py-3 text-left text-sm font-semibold transition ${selected ? "bg-white text-[#1a3a5c]" : "text-white/75 hover:bg-white/10 hover:text-white"}`}>
                      <Icon size={18} />{item.label}
                    </button>
                  );
                })}
              </nav>
            </aside>
          </div>
        )}

        <main className="flex-1">
          <header className="sticky top-0 z-40 border-b border-[#E2E8F0] bg-white px-5 py-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <button onClick={() => setMenuOpen(true)} className="lg:hidden rounded-xl bg-[#F7F8FA] p-2 text-[#1a3a5c]">
                  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round">
                    <line x1="3" y1="6" x2="21" y2="6" />
                    <line x1="3" y1="12" x2="21" y2="12" />
                    <line x1="3" y1="18" x2="21" y2="18" />
                  </svg>
                </button>
                <div>
                  <h1 className="text-2xl font-black">ARKA</h1>
                  <p className="text-sm text-[#718096] hidden sm:block">Professional Kitchen Management System</p>
                </div>
              </div>
              <button className="rounded-xl bg-[#1a3a5c] px-4 py-2 text-sm font-bold text-white">Close Week</button>
            </div>
          </header>
          <section className="p-5">{children}</section>
        </main>
      </div>
    </div>
  );
}

function Card({ children }: { children: React.ReactNode }) {
  return <div className="rounded-2xl border border-[#E2E8F0] bg-white p-5 shadow-sm">{children}</div>;
}

function KpiCard({ title, value, helper }: { title: string; value: string | number; helper: string }) {
  return (
    <Card>
      <p className="text-sm font-semibold uppercase tracking-wide text-[#718096]">{title}</p>
      <p className="mt-2 text-3xl font-black text-[#2D3748]">{value}</p>
      <p className="mt-2 text-sm text-[#718096]">{helper}</p>
    </Card>
  );
}

function Dashboard({ products, productionEntries, wasteEntries, salesEntries }: { products: Product[]; productionEntries: ProductionEntry[]; wasteEntries: WasteEntry[]; salesEntries: SaleEntry[] }) {
  const metrics = useMemo(() => {
    const currentStockValue = products.reduce((sum, item) => sum + item.realStock * item.price, 0);
    const theoreticalStockValue = products.reduce((sum, item) => sum + item.theoreticalStock * item.price, 0);
    const belowPar = products.filter((item) => item.realStock <= item.parLevel).length;
    const totalWaste = wasteEntries.reduce((sum, e) => sum + e.estimatedCost, 0);
    const totalSales = salesEntries.reduce((sum, e) => sum + e.totalRevenue, 0);
    const todayStr = new Date().toISOString().split("T")[0];
    const todayProduction = productionEntries.filter((e) => e.date === todayStr).length;
    return { currentStockValue, theoreticalStockValue, belowPar, totalWaste, todayProduction, totalSales };
  }, [products, productionEntries, wasteEntries, salesEntries]);

  return (
    <>
      <div className="mb-5">
        <h2 className="text-3xl font-black">Dashboard</h2>
        <p className="text-[#718096]">Real-time overview of stock, production and kitchen control.</p>
      </div>
      <div className="grid gap-4 md:grid-cols-3 xl:grid-cols-6">
        <KpiCard title="Current Stock" value={money(metrics.currentStockValue)} helper="Real stock value" />
        <KpiCard title="Theoretical Stock" value={money(metrics.theoreticalStockValue)} helper="Expected" />
        <KpiCard title="Sales" value={money(metrics.totalSales)} helper="This week" />
        <KpiCard title="Waste" value={money(metrics.totalWaste)} helper="This week" />
        <KpiCard title="Production" value={metrics.todayProduction} helper="Entries today" />
        <KpiCard title="Below PAR" value={metrics.belowPar} helper="Needs review" />
      </div>
      <div className="mt-5 grid gap-4 xl:grid-cols-2">
        <Card>
          <h3 className="mb-4 text-xl font-black">Stock Alerts</h3>
          <div className="grid gap-3">
            {products.filter(p => p.realStock <= p.parLevel).slice(0, 5).map((item) => (
              <div key={item.id} className="flex items-center justify-between rounded-xl bg-[#F7F8FA] p-3">
                <div><b>{item.name}</b><p className="text-sm text-[#718096]">{item.category} · {item.unit}</p></div>
                <span className="rounded-full bg-[#FDE8C8] px-3 py-1 text-xs font-bold text-[#E07B00]">Review</span>
              </div>
            ))}
            {products.filter(p => p.realStock > p.parLevel).slice(0, 2).map((item) => (
              <div key={item.id} className="flex items-center justify-between rounded-xl bg-[#F7F8FA] p-3">
                <div><b>{item.name}</b><p className="text-sm text-[#718096]">{item.category} · {item.unit}</p></div>
                <span className="rounded-full bg-[#D4EDE0] px-3 py-1 text-xs font-bold text-[#4A7C59]">OK</span>
              </div>
            ))}
          </div>
        </Card>
        <Card>
          <h3 className="mb-4 text-xl font-black">Latest Activity</h3>
          <div className="grid gap-2">
            {[
              ...productionEntries.slice(0, 2).map(e => ({ type: "production", text: `${e.recipeName} × ${e.portions}`, date: e.date })),
              ...wasteEntries.slice(0, 2).map(e => ({ type: "waste", text: `${e.refName} — ${e.reason}`, date: e.date })),
              ...salesEntries.slice(0, 2).map(e => ({ type: "sales", text: `Sales: ${money(e.totalRevenue)}`, date: e.date })),
            ].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 6).map((item, i) => (
              <div key={i} className="flex items-center gap-3 rounded-xl bg-[#F7F8FA] p-3">
                <span className={`rounded-full px-2 py-0.5 text-xs font-bold ${item.type === "production" ? "bg-[#E6F1FB] text-[#185FA5]" : item.type === "sales" ? "bg-[#D4EDE0] text-[#4A7C59]" : "bg-[#FDE8C8] text-[#E07B00]"}`}>
                  {item.type === "production" ? "Prod" : item.type === "sales" ? "Sales" : "Waste"}
                </span>
                <p className="text-sm flex-1">{item.text}</p>
                <p className="text-xs text-[#718096]">{item.date}</p>
              </div>
            ))}
            {productionEntries.length === 0 && wasteEntries.length === 0 && salesEntries.length === 0 && (
              <div className="rounded-xl bg-[#F7F8FA] p-4 text-[#718096]">No activity yet.</div>
            )}
          </div>
        </Card>
      </div>
    </>
  );
}

function Stocktake({ products, setProducts }: { products: Product[]; setProducts: React.Dispatch<React.SetStateAction<Product[]>> }) {
  const [query, setQuery] = useState("");
  const [importMsg, setImportMsg] = useState("");
  const [editingId, setEditingId] = useState<number | null>(null);
  const [showAddForm, setShowAddForm] = useState(false);
  const [saveMsg, setSaveMsg] = useState("");
  const [aiMsg, setAiMsg] = useState("");
  const [aiLoading, setAiLoading] = useState(false);
  const [aiPreview, setAiPreview] = useState<AiProduct[]>([]);
  const [newProduct, setNewProduct] = useState<Omit<Product, "id">>({ name: "", category: "Meat", unit: "kg", price: 0, realStock: 0, theoreticalStock: 0, parLevel: 0 });
  const fileRef = useRef<HTMLInputElement>(null);
  const aiFileRef = useRef<HTMLInputElement>(null);

  const filtered = products.filter((item) => item.name.toLowerCase().includes(query.toLowerCase()));

  const updateQty = (id: number, value: number) => setProducts((items) => items.map((item) => item.id === id ? { ...item, realStock: Math.max(0, value) } : item));
  const quickUpdate = (id: number, delta: number) => setProducts((items) => items.map((item) => item.id === id ? { ...item, realStock: Math.max(0, Number(item.realStock) + delta) } : item));
  const updateField = (id: number, field: keyof Product, value: string | number) => setProducts((items) => items.map((item) => item.id === id ? { ...item, [field]: value } : item));
  const deleteProduct = (id: number) => { setProducts((items) => items.filter((item) => item.id !== id)); setEditingId(null); };
  const saveProduct = () => { setEditingId(null); setSaveMsg("✅ Saved"); setTimeout(() => setSaveMsg(""), 2000); };

  const addProduct = () => {
    if (!newProduct.name.trim()) return;
    setProducts((prev) => [...prev, { ...newProduct, id: Date.now(), theoreticalStock: newProduct.realStock }]);
    setNewProduct({ name: "", category: "Meat", unit: "kg", price: 0, realStock: 0, theoreticalStock: 0, parLevel: 0 });
    setShowAddForm(false);
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportMsg("Reading file...");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        let headerRow = -1;
        for (let i = 0; i < rows.length; i++) { if (String(rows[i][0]).toLowerCase().includes("item")) { headerRow = i; break; } }
        if (headerRow === -1) { setImportMsg("❌ Could not find 'Item' column."); return; }
        const headers = rows[headerRow];
        let lastCountCol = -1;
        for (let c = 0; c < headers.length; c++) { if (String(headers[c]).toUpperCase().includes("COUNT")) lastCountCol = c; }
        if (lastCountCol === -1) { setImportMsg("❌ Could not find COUNT column."); return; }
        let updated = 0; let added = 0;
        const newProducts: Product[] = [...products];
        for (let i = headerRow + 1; i < rows.length; i++) {
          const row = rows[i];
          const name = String(row[0] || "").trim();
          const unit = String(row[1] || "").trim();
          const par = parseFloat(String(row[2])) || 0;
          const price = parseFloat(String(row[3])) || 0;
          const count = parseFloat(String(row[lastCountCol]));
          if (!name || isNaN(count) || ["🥩", "❄️", "🧀", "🥗", "🫙", "🍷", "📦"].some(e => name.startsWith(e))) continue;
          const existingIdx = newProducts.findIndex(p => p.name.toLowerCase() === name.toLowerCase());
          if (existingIdx >= 0) { newProducts[existingIdx] = { ...newProducts[existingIdx], realStock: count, theoreticalStock: count, parLevel: par || newProducts[existingIdx].parLevel, price: price || newProducts[existingIdx].price }; updated++; }
          else { newProducts.push({ id: Date.now() + i, name, category: "Imported", unit, price, realStock: count, theoreticalStock: count, parLevel: par }); added++; }
        }
        setProducts(newProducts);
        setImportMsg(`✅ Done! ${updated} updated, ${added} new products added.`);
        if (fileRef.current) fileRef.current.value = "";
      } catch { setImportMsg("❌ Error reading file."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleAiImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setAiLoading(true);
    setAiMsg("Reading with AI...");
    setAiPreview([]);
    try {
      const toBase64 = (f: File): Promise<string> => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve((reader.result as string).split(",")[1]);
        reader.onerror = reject;
        reader.readAsDataURL(f);
      });
      const base64 = await toBase64(file);
      const mediaType = (file.type || "image/jpeg") as "image/jpeg" | "image/png" | "image/webp" | "image/gif";
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": import.meta.env.VITE_ANTHROPIC_API_KEY,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify({
          model: "claude-opus-4-5",
          max_tokens: 1024,
          messages: [{
            role: "user",
            content: [
              { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
              { type: "text", text: `This is a delivery note or invoice from a restaurant supplier. Extract all products with their quantities, units and prices if visible. Return ONLY a JSON array like this, no other text:\n[{"name":"Product Name","quantity":5,"unit":"kg","price":12.50}, ...]\nIf price is not visible use 0. Be precise with quantities and units.` },
            ],
          }],
        }),
      });
      const data = await response.json();
      const text = data.content?.[0]?.text || "";
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      setAiPreview(parsed);
      setAiMsg(`✅ AI found ${parsed.length} products. Review and confirm below.`);
    } catch {
      setAiMsg("❌ Could not read the file. Try a clearer photo.");
    } finally {
      setAiLoading(false);
      if (aiFileRef.current) aiFileRef.current.value = "";
    }
  };

  const confirmAiImport = () => {
    const newProducts: Product[] = [...products];
    let updated = 0; let added = 0;
    aiPreview.forEach((item, i) => {
      const existingIdx = newProducts.findIndex(p => p.name.toLowerCase() === item.name.toLowerCase());
      if (existingIdx >= 0) {
        newProducts[existingIdx] = { ...newProducts[existingIdx], realStock: item.quantity, theoreticalStock: item.quantity, price: item.price || newProducts[existingIdx].price };
        updated++;
      } else {
        newProducts.push({ id: Date.now() + i, name: item.name, category: "Imported", unit: item.unit, price: item.price, realStock: item.quantity, theoreticalStock: item.quantity, parLevel: 0 });
        added++;
      }
    });
    setProducts(newProducts);
    setAiPreview([]);
    setAiMsg(`✅ Done! ${updated} updated, ${added} new products added to stock.`);
  };

  return (
    <>
      <div className="mb-5 flex items-center justify-between">
        <div><h2 className="text-3xl font-black">Stocktake</h2><p className="text-[#718096]">Count, edit and review current kitchen stock.</p></div>
        <button onClick={() => setShowAddForm(true)} className="flex items-center gap-2 rounded-xl bg-[#1a3a5c] px-5 py-3 font-bold text-white"><Plus size={18} /> Add Product</button>
      </div>

      {showAddForm && (
        <Card>
          <h3 className="mb-4 font-black text-lg">New Product</h3>
          <div className="grid gap-3 md:grid-cols-3 lg:grid-cols-6">
            <div className="lg:col-span-2"><p className="mb-1 text-xs font-semibold text-[#718096]">NAME</p><input value={newProduct.name} onChange={(e) => setNewProduct(p => ({ ...p, name: e.target.value }))} placeholder="Product name" className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">CATEGORY</p><select value={newProduct.category} onChange={(e) => setNewProduct(p => ({ ...p, category: e.target.value }))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none">{CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">UNIT</p><select value={newProduct.unit} onChange={(e) => setNewProduct(p => ({ ...p, unit: e.target.value }))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none">{UNITS.map(u => <option key={u}>{u}</option>)}</select></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">PRICE £</p><input type="number" value={newProduct.price} onChange={(e) => setNewProduct(p => ({ ...p, price: Number(e.target.value) }))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">PAR</p><input type="number" value={newProduct.parLevel} onChange={(e) => setNewProduct(p => ({ ...p, parLevel: Number(e.target.value) }))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
          </div>
          <div className="mt-4 flex gap-3">
            <button onClick={addProduct} className="rounded-xl bg-[#1a3a5c] px-5 py-2 font-bold text-white">Save</button>
            <button onClick={() => setShowAddForm(false)} className="rounded-xl border border-[#E2E8F0] px-5 py-2 font-bold text-[#718096]">Cancel</button>
          </div>
        </Card>
      )}

      <Card>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="relative flex-1">
            <Search size={18} className="absolute left-4 top-3.5 text-[#718096]" />
            <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search product" className="h-12 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] pl-11 pr-4 outline-none" />
          </div>
          <label className="flex h-12 cursor-pointer items-center gap-2 rounded-xl border-2 border-dashed border-[#1a3a5c] px-5 font-bold text-[#1a3a5c] hover:bg-[#1a3a5c] hover:text-white transition">
            <Upload size={18} /> Import Excel
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleImport} />
          </label>
          <label className="flex h-12 cursor-pointer items-center gap-2 rounded-xl border-2 border-dashed border-[#E07B00] px-5 font-bold text-[#E07B00] hover:bg-[#E07B00] hover:text-white transition">
            {aiLoading ? "🤖 Reading..." : "🤖 AI Read"}
            <input ref={aiFileRef} type="file" accept="image/*,.pdf" className="hidden" onChange={handleAiImport} />
          </label>
        </div>
        {importMsg && <div className={`mt-3 rounded-xl px-4 py-3 text-sm font-semibold ${importMsg.startsWith("✅") ? "bg-[#D4EDE0] text-[#4A7C59]" : importMsg.startsWith("❌") ? "bg-red-50 text-red-600" : "bg-[#E6F1FB] text-[#185FA5]"}`}>{importMsg}</div>}
        {saveMsg && <div className="mt-3 rounded-xl px-4 py-3 text-sm font-semibold bg-[#D4EDE0] text-[#4A7C59]">{saveMsg}</div>}
        {(aiMsg || aiLoading) && (
          <div className={`mt-3 rounded-xl px-4 py-3 text-sm font-semibold ${aiMsg.startsWith("✅") ? "bg-[#D4EDE0] text-[#4A7C59]" : aiMsg.startsWith("❌") ? "bg-red-50 text-red-600" : "bg-[#E6F1FB] text-[#185FA5]"}`}>
            {aiLoading ? "🤖 Reading with AI, please wait..." : aiMsg}
          </div>
        )}
        {aiPreview.length > 0 && (
          <div className="mt-3">
            <p className="mb-2 text-sm font-black text-[#718096]">AI PREVIEW — Review before confirming:</p>
            <div className="grid gap-2 mb-3">
              {aiPreview.map((item, i) => (
                <div key={i} className="flex items-center justify-between rounded-xl bg-[#FDE8C8]/50 border border-[#E07B00]/20 p-3">
                  <div><b>{item.name}</b><p className="text-xs text-[#718096]">{item.unit}</p></div>
                  <div className="text-right"><p className="font-black">{item.quantity} {item.unit}</p><p className="text-xs text-[#718096]">£{item.price}/unit</p></div>
                </div>
              ))}
            </div>
            <div className="flex gap-3">
              <button onClick={confirmAiImport} className="rounded-xl bg-[#E07B00] px-5 py-2 font-bold text-white">Confirm & Update Stock</button>
              <button onClick={() => { setAiPreview([]); setAiMsg(""); }} className="rounded-xl border border-[#E2E8F0] px-5 py-2 font-bold text-[#718096]">Cancel</button>
            </div>
          </div>
        )}
      </Card>

      <div className="mt-4 grid gap-3">
        {filtered.map((item) => (
          <Card key={item.id}>
            {editingId === item.id ? (
              <div>
                <div className="grid gap-3 md:grid-cols-3 lg:grid-cols-6 mb-4">
                  <div className="lg:col-span-2"><p className="mb-1 text-xs font-semibold text-[#718096]">NAME</p><input value={item.name} onChange={(e) => updateField(item.id, "name", e.target.value)} className="h-10 w-full rounded-xl border border-[#1a3a5c] bg-white px-4 outline-none text-sm font-semibold" /></div>
                  <div><p className="mb-1 text-xs font-semibold text-[#718096]">CATEGORY</p><select value={item.category} onChange={(e) => updateField(item.id, "category", e.target.value)} className="h-10 w-full rounded-xl border border-[#1a3a5c] bg-white px-3 outline-none text-sm">{CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></div>
                  <div><p className="mb-1 text-xs font-semibold text-[#718096]">UNIT</p><select value={item.unit} onChange={(e) => updateField(item.id, "unit", e.target.value)} className="h-10 w-full rounded-xl border border-[#1a3a5c] bg-white px-3 outline-none text-sm">{UNITS.map(u => <option key={u}>{u}</option>)}</select></div>
                  <div><p className="mb-1 text-xs font-semibold text-[#718096]">PRICE £</p><input type="number" value={item.price} onChange={(e) => updateField(item.id, "price", Number(e.target.value))} className="h-10 w-full rounded-xl border border-[#1a3a5c] bg-white px-3 outline-none text-sm" /></div>
                  <div><p className="mb-1 text-xs font-semibold text-[#718096]">PAR</p><input type="number" value={item.parLevel} onChange={(e) => updateField(item.id, "parLevel", Number(e.target.value))} className="h-10 w-full rounded-xl border border-[#1a3a5c] bg-white px-3 outline-none text-sm" /></div>
                </div>
                <div className="flex gap-3">
                  <button onClick={saveProduct} className="flex items-center gap-2 rounded-xl bg-[#1a3a5c] px-5 py-2 text-sm font-bold text-white"><Save size={14} /> Save changes</button>
                  <button onClick={() => deleteProduct(item.id)} className="rounded-xl bg-red-50 px-5 py-2 text-sm font-bold text-red-500 hover:bg-red-100">Delete</button>
                  <button onClick={() => setEditingId(null)} className="rounded-xl border border-[#E2E8F0] px-5 py-2 text-sm font-bold text-[#718096]">Cancel</button>
                </div>
              </div>
            ) : (
              <div className="grid gap-4 lg:grid-cols-[1fr_180px] lg:items-center">
                <div>
                  <div className="flex items-start justify-between gap-3">
                    <div><h3 className="font-black">{item.name}</h3><p className="text-sm text-[#718096]">{item.category} · {money(item.price)} / {item.unit}</p></div>
                    <div className="flex items-center gap-2">
                      <span className={`rounded-full px-3 py-1 text-xs font-bold ${item.realStock <= item.parLevel ? "bg-[#FDE8C8] text-[#E07B00]" : "bg-[#E6F1FB] text-[#185FA5]"}`}>PAR {item.parLevel}</span>
                      <button onClick={() => setEditingId(item.id)} className="rounded-lg bg-[#F7F8FA] px-3 py-1 text-xs font-bold text-[#718096] hover:bg-[#E2E8F0]">Edit</button>
                    </div>
                  </div>
                  <div className="mt-3 grid max-w-md grid-cols-3 gap-2">
                    <div className="rounded-xl bg-[#F7F8FA] p-2"><p className="text-xs text-[#718096]">Real</p><b>{item.realStock}</b></div>
                    <div className="rounded-xl bg-[#F7F8FA] p-2"><p className="text-xs text-[#718096]">Theoretical</p><b>{item.theoreticalStock}</b></div>
                    <div className="rounded-xl bg-[#F7F8FA] p-2"><p className="text-xs text-[#718096]">PAR</p><b>{item.parLevel}</b></div>
                  </div>
                </div>
                <div className="flex items-center justify-end gap-2">
                  <button onClick={() => quickUpdate(item.id, -1)} className="h-11 w-11 rounded-xl border border-[#1a3a5c] text-[#1a3a5c]">-</button>
                  <input type="number" value={item.realStock} onChange={(e) => updateQty(item.id, Number(e.target.value))} className="h-12 w-24 rounded-xl border border-[#E2E8F0] text-center font-black outline-none" />
                  <button onClick={() => quickUpdate(item.id, 1)} className="h-11 w-11 rounded-xl bg-[#1a3a5c] text-white">+</button>
                </div>
              </div>
            )}
          </Card>
        ))}
      </div>
    </>
  );
}

function Recipes({ products, recipes, setRecipes }: { products: Product[]; recipes: Recipe[]; setRecipes: React.Dispatch<React.SetStateAction<Recipe[]>> }) {
  const [selected, setSelected] = useState<Recipe | null>(recipes[0] ?? null);
  const [showForm, setShowForm] = useState(false);
  const [tab, setTab] = useState<"ingredients" | "procedure" | "info">("ingredients");
  const [newName, setNewName] = useState("");
  const [newCat, setNewCat] = useState("Mains");
  const [newPortions, setNewPortions] = useState(1);
  const [newPrepTime, setNewPrepTime] = useState(30);
  const [newIngProduct, setNewIngProduct] = useState(products[0]?.id ?? 1);
  const [newIngQty, setNewIngQty] = useState(1);
  const [newIngUnit, setNewIngUnit] = useState("kg");
  const [newStep, setNewStep] = useState("");
  const categories = ["Starters", "Mains", "Desserts", "Sauces", "Bases"];

  const addRecipe = () => {
    if (!newName.trim()) return;
    const r: Recipe = { id: Date.now(), name: newName, category: newCat, portions: newPortions, prepTime: newPrepTime, allergens: [], procedure: [], photo: null, ingredients: [] };
    setRecipes((prev) => [...prev, r]); setSelected(r); setNewName(""); setShowForm(false);
  };
  const deleteRecipe = (id: number) => { setRecipes((prev) => prev.filter(r => r.id !== id)); setSelected(null); };
  const updateSelected = (changes: Partial<Recipe>) => {
    if (!selected) return;
    const updated = { ...selected, ...changes };
    setSelected(updated); setRecipes((prev) => prev.map((r) => r.id === updated.id ? updated : r));
  };
  const addIngredient = () => { if (!selected) return; updateSelected({ ingredients: [...selected.ingredients, { id: Date.now(), productId: newIngProduct, quantity: newIngQty, unit: newIngUnit }] }); };
  const removeIngredient = (id: number) => { if (!selected) return; updateSelected({ ingredients: selected.ingredients.filter((i) => i.id !== id) }); };
  const addStep = () => { if (!selected || !newStep.trim()) return; updateSelected({ procedure: [...selected.procedure, newStep] }); setNewStep(""); };
  const removeStep = (idx: number) => { if (!selected) return; updateSelected({ procedure: selected.procedure.filter((_, i) => i !== idx) }); };
  const toggleAllergen = (a: string) => { if (!selected) return; const has = selected.allergens.includes(a); updateSelected({ allergens: has ? selected.allergens.filter((x) => x !== a) : [...selected.allergens, a] }); };
  const handlePhoto = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file || !selected) return;
    const reader = new FileReader();
    reader.onload = () => updateSelected({ photo: reader.result as string });
    reader.readAsDataURL(file);
  };
  const totalCost = selected?.ingredients.reduce((sum, ing) => { const p = products.find((p) => p.id === ing.productId); return sum + (p?.price ?? 0) * ing.quantity; }, 0) ?? 0;
  const costPerPortion = selected ? totalCost / selected.portions : 0;

  return (
    <>
      <div className="mb-5 flex items-center justify-between">
        <div><h2 className="text-3xl font-black">Recipes</h2><p className="text-[#718096]">Technical sheets, ingredients, procedure and costs.</p></div>
        <button onClick={() => setShowForm(true)} className="flex items-center gap-2 rounded-xl bg-[#1a3a5c] px-5 py-3 font-bold text-white"><Plus size={18} /> New Recipe</button>
      </div>
      {showForm && (
        <Card>
          <h3 className="mb-4 font-black text-lg">New Recipe</h3>
          <div className="grid gap-3 md:grid-cols-4">
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">NAME</p><input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="Recipe name" className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">CATEGORY</p><select value={newCat} onChange={(e) => setNewCat(e.target.value)} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none">{categories.map((c) => <option key={c}>{c}</option>)}</select></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">PORTIONS</p><input type="number" value={newPortions} onChange={(e) => setNewPortions(Number(e.target.value))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">PREP TIME (min)</p><input type="number" value={newPrepTime} onChange={(e) => setNewPrepTime(Number(e.target.value))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
          </div>
          <div className="mt-4 flex gap-3">
            <button onClick={addRecipe} className="rounded-xl bg-[#1a3a5c] px-5 py-2 font-bold text-white">Save</button>
            <button onClick={() => setShowForm(false)} className="rounded-xl border border-[#E2E8F0] px-5 py-2 font-bold text-[#718096]">Cancel</button>
          </div>
        </Card>
      )}
      <div className="mt-4 grid gap-4 lg:grid-cols-[260px_1fr]">
        <div className="grid gap-2 content-start">
          {recipes.map((r) => (
            <button key={r.id} onClick={() => { setSelected(r); setTab("ingredients"); }}
              className={`rounded-2xl border p-4 text-left transition ${selected?.id === r.id ? "border-[#1a3a5c] bg-[#1a3a5c] text-white" : "border-[#E2E8F0] bg-white hover:border-[#1a3a5c]"}`}>
              <p className="font-black">{r.name}</p>
              <p className={`text-sm ${selected?.id === r.id ? "text-white/70" : "text-[#718096]"}`}>{r.category} · {r.portions} pax · {r.prepTime}min</p>
            </button>
          ))}
        </div>
        {selected ? (
          <div className="grid gap-4">
            <Card>
              <div className="flex gap-4">
                <div className="relative h-28 w-28 flex-shrink-0 overflow-hidden rounded-xl bg-[#F7F8FA]">
                  {selected.photo ? <img src={selected.photo} alt={selected.name} className="h-full w-full object-cover" /> : <div className="flex h-full w-full items-center justify-center text-[#718096]"><Utensils size={32} /></div>}
                  <label className="absolute inset-0 flex cursor-pointer items-end justify-center bg-black/30 opacity-0 hover:opacity-100 transition">
                    <span className="mb-2 text-xs font-bold text-white">Upload photo</span>
                    <input type="file" accept="image/*" className="hidden" onChange={handlePhoto} />
                  </label>
                </div>
                <div className="flex-1">
                  <div className="flex items-start justify-between">
                    <div><h3 className="text-xl font-black">{selected.name}</h3><p className="text-sm text-[#718096]">{selected.category}</p></div>
                    <button onClick={() => deleteRecipe(selected.id)} className="rounded-lg bg-red-50 px-3 py-1.5 text-xs font-bold text-red-500 hover:bg-red-100">Delete recipe</button>
                  </div>
                  <div className="mt-3 flex flex-wrap gap-2">
                    <span className="rounded-full bg-[#E6F1FB] px-3 py-1 text-xs font-bold text-[#185FA5]">{selected.portions} portions</span>
                    <span className="rounded-full bg-[#F7F8FA] px-3 py-1 text-xs font-bold text-[#718096]">⏱ {selected.prepTime} min</span>
                    <span className="rounded-full bg-[#D4EDE0] px-3 py-1 text-xs font-bold text-[#4A7C59]">Cost/pax: {money(costPerPortion)}</span>
                    <span className="rounded-full bg-[#FDE8C8] px-3 py-1 text-xs font-bold text-[#E07B00]">Total: {money(totalCost)}</span>
                  </div>
                  {selected.allergens.length > 0 && <div className="mt-2 flex flex-wrap gap-1">{selected.allergens.map((a) => <span key={a} className="rounded-full bg-red-100 px-2 py-0.5 text-xs font-bold text-red-600">{a}</span>)}</div>}
                </div>
              </div>
            </Card>
            <div className="flex gap-2">
              {(["ingredients", "procedure", "info"] as const).map((t) => (
                <button key={t} onClick={() => setTab(t)} className={`rounded-xl px-4 py-2 text-sm font-bold capitalize transition ${tab === t ? "bg-[#1a3a5c] text-white" : "bg-white border border-[#E2E8F0] text-[#718096] hover:border-[#1a3a5c]"}`}>{t}</button>
              ))}
            </div>
            {tab === "ingredients" && (
              <Card>
                <h4 className="mb-3 font-black">Ingredients</h4>
                <div className="grid gap-2 mb-4">
                  {selected.ingredients.map((ing) => {
                    const product = products.find((p) => p.id === ing.productId);
                    return (
                      <div key={ing.id} className="flex items-center justify-between rounded-xl bg-[#F7F8FA] p-3">
                        <div><b>{product?.name ?? "Unknown"}</b><p className="text-sm text-[#718096]">{product?.category}</p></div>
                        <div className="flex items-center gap-3">
                          <div className="text-right"><p className="font-black">{ing.quantity} {ing.unit}</p><p className="text-sm text-[#718096]">{money((product?.price ?? 0) * ing.quantity)}</p></div>
                          <button onClick={() => removeIngredient(ing.id)} className="rounded-lg bg-red-50 p-2 text-red-400 hover:bg-red-100"><Trash2 size={14} /></button>
                        </div>
                      </div>
                    );
                  })}
                </div>
                <div className="grid gap-2 md:grid-cols-4 items-end border-t border-[#E2E8F0] pt-4">
                  <div><p className="mb-1 text-xs font-semibold text-[#718096]">PRODUCT</p>
                    <select value={newIngProduct} onChange={(e) => setNewIngProduct(Number(e.target.value))} className="h-10 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm">
                      {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  </div>
                  <div><p className="mb-1 text-xs font-semibold text-[#718096]">QTY</p><input type="number" value={newIngQty} onChange={(e) => setNewIngQty(Number(e.target.value))} className="h-10 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm" /></div>
                  <div><p className="mb-1 text-xs font-semibold text-[#718096]">UNIT</p>
                    <select value={newIngUnit} onChange={(e) => setNewIngUnit(e.target.value)} className="h-10 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm">
                      {["kg", "g", "l", "ml", "unit", "tbsp", "tsp"].map((u) => <option key={u}>{u}</option>)}
                    </select>
                  </div>
                  <button onClick={addIngredient} className="h-10 rounded-xl bg-[#1a3a5c] px-4 font-bold text-white text-sm flex items-center gap-2 justify-center"><Plus size={16} /> Add</button>
                </div>
              </Card>
            )}
            {tab === "procedure" && (
              <Card>
                <h4 className="mb-3 font-black">Procedure</h4>
                <div className="grid gap-2 mb-4">
                  {selected.procedure.length === 0 ? <div className="rounded-xl bg-[#F7F8FA] p-4 text-center text-[#718096] text-sm">No steps yet.</div> :
                    selected.procedure.map((step, idx) => (
                      <div key={idx} className="flex items-start gap-3 rounded-xl bg-[#F7F8FA] p-3">
                        <span className="flex h-7 w-7 flex-shrink-0 items-center justify-center rounded-full bg-[#1a3a5c] text-xs font-black text-white">{idx + 1}</span>
                        <p className="flex-1 text-sm pt-1">{step}</p>
                        <button onClick={() => removeStep(idx)} className="rounded-lg bg-red-50 p-1.5 text-red-400 hover:bg-red-100"><Trash2 size={13} /></button>
                      </div>
                    ))}
                </div>
                <div className="flex gap-2 border-t border-[#E2E8F0] pt-4">
                  <input value={newStep} onChange={(e) => setNewStep(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addStep()} placeholder="Write a step and press Enter or Add" className="h-10 flex-1 rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none text-sm" />
                  <button onClick={addStep} className="h-10 rounded-xl bg-[#1a3a5c] px-4 font-bold text-white text-sm flex items-center gap-2"><Plus size={16} /> Add</button>
                </div>
              </Card>
            )}
            {tab === "info" && (
              <Card>
                <h4 className="mb-3 font-black">Allergens & Info</h4>
                <p className="mb-2 text-xs font-semibold text-[#718096]">ALLERGENS</p>
                <div className="flex flex-wrap gap-2 mb-6">
                  {ALLERGENS.map((a) => <button key={a} onClick={() => toggleAllergen(a)} className={`rounded-full px-3 py-1 text-xs font-bold transition ${selected.allergens.includes(a) ? "bg-red-500 text-white" : "bg-[#F7F8FA] text-[#718096] hover:bg-red-50"}`}>{a}</button>)}
                </div>
                <p className="mb-2 text-xs font-semibold text-[#718096]">PREP TIME</p>
                <div className="flex items-center gap-3">
                  <input type="number" value={selected.prepTime} onChange={(e) => updateSelected({ prepTime: Number(e.target.value) })} className="h-10 w-24 rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm" />
                  <span className="text-sm text-[#718096]">minutes</span>
                </div>
              </Card>
            )}
          </div>
        ) : (
          <div className="flex h-48 items-center justify-center rounded-2xl border-2 border-dashed border-[#E2E8F0] text-[#718096]">Select a recipe to see its details</div>
        )}
      </div>
    </>
  );
}

function Production({ entries, setEntries, recipes }: { entries: ProductionEntry[]; setEntries: React.Dispatch<React.SetStateAction<ProductionEntry[]>>; recipes: Recipe[] }) {
  const [showForm, setShowForm] = useState(false);
  const [recipeId, setRecipeId] = useState(recipes[0]?.id ?? 1);
  const [portions, setPortions] = useState(1);
  const [notes, setNotes] = useState("");
  const [date, setDate] = useState(new Date().toISOString().split("T")[0]);

  const addEntry = () => {
    const recipe = recipes.find((r) => r.id === recipeId);
    if (!recipe) return;
    setEntries((prev) => [{ id: Date.now(), date, recipeId, recipeName: recipe.name, portions, notes }, ...prev]);
    setPortions(1); setNotes(""); setShowForm(false);
  };
  const deleteEntry = (id: number) => setEntries((prev) => prev.filter(e => e.id !== id));
  const grouped = entries.reduce((acc, e) => { if (!acc[e.date]) acc[e.date] = []; acc[e.date].push(e); return acc; }, {} as Record<string, ProductionEntry[]>);

  return (
    <>
      <div className="mb-5 flex items-center justify-between">
        <div><h2 className="text-3xl font-black">Production</h2><p className="text-[#718096]">Log daily production. History saved for weekly close.</p></div>
        <button onClick={() => setShowForm(true)} className="flex items-center gap-2 rounded-xl bg-[#1a3a5c] px-5 py-3 font-bold text-white"><Plus size={18} /> Log Production</button>
      </div>
      <div className="grid gap-4 md:grid-cols-3 mb-5">
        <KpiCard title="Total Entries" value={entries.length} helper="This week" />
        <KpiCard title="Total Portions" value={entries.reduce((s, e) => s + e.portions, 0)} helper="Produced this week" />
        <KpiCard title="Recipes Used" value={new Set(entries.map(e => e.recipeId)).size} helper="Different recipes" />
      </div>
      {showForm && (
        <Card>
          <h3 className="mb-4 font-black text-lg">Log Production</h3>
          <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-4">
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">DATE</p><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">RECIPE</p>
              <select value={recipeId} onChange={(e) => setRecipeId(Number(e.target.value))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none">
                {recipes.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
              </select>
            </div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">PORTIONS</p><input type="number" value={portions} onChange={(e) => setPortions(Number(e.target.value))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">NOTES</p><input value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Optional" className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
          </div>
          <div className="mt-4 flex gap-3">
            <button onClick={addEntry} className="rounded-xl bg-[#1a3a5c] px-5 py-2 font-bold text-white">Save</button>
            <button onClick={() => setShowForm(false)} className="rounded-xl border border-[#E2E8F0] px-5 py-2 font-bold text-[#718096]">Cancel</button>
          </div>
        </Card>
      )}
      <div className="mt-4 grid gap-5">
        {Object.keys(grouped).length === 0 ? <Card><div className="py-8 text-center text-[#718096]">No production entries yet.</div></Card> :
          Object.keys(grouped).sort((a, b) => b.localeCompare(a)).map(date => (
            <div key={date}>
              <p className="mb-2 text-sm font-black text-[#718096] uppercase tracking-wide">{date}</p>
              <div className="grid gap-3">
                {grouped[date].map((e) => (
                  <Card key={e.id}>
                    <div className="flex items-center justify-between">
                      <div><h3 className="font-black">{e.recipeName}</h3><p className="text-sm text-[#718096]">{e.notes || "No notes"}</p></div>
                      <div className="flex items-center gap-4">
                        <div className="text-right"><p className="text-2xl font-black text-[#1a3a5c]">{e.portions}</p><p className="text-xs text-[#718096]">portions</p></div>
                        <button onClick={() => deleteEntry(e.id)} className="rounded-lg bg-red-50 p-2 text-red-400 hover:bg-red-100"><Trash2 size={14} /></button>
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
            </div>
          ))}
      </div>
    </>
  );
}

function Waste({ products, entries, setEntries, productionEntries }: { products: Product[]; entries: WasteEntry[]; setEntries: React.Dispatch<React.SetStateAction<WasteEntry[]>>; productionEntries: ProductionEntry[] }) {
  const [showForm, setShowForm] = useState(false);
  const [wasteType, setWasteType] = useState<"product" | "production">("product");
  const [refId, setRefId] = useState(products[0]?.id ?? 1);
  const [quantity, setQuantity] = useState(1);
  const [unit, setUnit] = useState("kg");
  const [reason, setReason] = useState("Spoilage");
  const [estimatedCost, setEstimatedCost] = useState(0);
  const [notes, setNotes] = useState("");
  const [date, setDate] = useState(new Date().toISOString().split("T")[0]);

  const currentOptions = wasteType === "product"
    ? products.map(p => ({ id: p.id, name: p.name, unit: p.unit, price: p.price }))
    : productionEntries.map(e => ({ id: e.id, name: `${e.recipeName} (${e.date} · ${e.portions} portions)`, unit: "portions", price: 0 }));

  const handleTypeChange = (t: "product" | "production") => {
    setWasteType(t);
    if (t === "product" && products.length > 0) setRefId(products[0].id);
    if (t === "production" && productionEntries.length > 0) setRefId(productionEntries[0].id);
  };

  const handleRefChange = (id: number) => {
    setRefId(id);
    if (wasteType === "product") { const p = products.find(p => p.id === id); if (p) { setUnit(p.unit); setEstimatedCost(p.price * quantity); } }
  };

  const addEntry = () => {
    const ref = currentOptions.find(o => o.id === refId);
    if (!ref) return;
    setEntries((prev) => [{ id: Date.now(), date, type: wasteType, refId, refName: ref.name, quantity, unit, reason, estimatedCost, notes }, ...prev]);
    setQuantity(1); setNotes(""); setEstimatedCost(0); setShowForm(false);
  };

  const deleteEntry = (id: number) => setEntries((prev) => prev.filter(e => e.id !== id));
  const totalWaste = entries.reduce((s, e) => s + e.estimatedCost, 0);
  const productWaste = entries.filter(e => e.type === "product").reduce((s, e) => s + e.estimatedCost, 0);
  const productionWaste = entries.filter(e => e.type === "production").reduce((s, e) => s + e.estimatedCost, 0);
  const grouped = entries.reduce((acc, e) => { if (!acc[e.date]) acc[e.date] = []; acc[e.date].push(e); return acc; }, {} as Record<string, WasteEntry[]>);

  return (
    <>
      <div className="mb-5 flex items-center justify-between">
        <div><h2 className="text-3xl font-black">Waste</h2><p className="text-[#718096]">Log losses from stock or production. Saved for weekly close.</p></div>
        <button onClick={() => setShowForm(true)} className="flex items-center gap-2 rounded-xl bg-[#1a3a5c] px-5 py-3 font-bold text-white"><Plus size={18} /> Log Waste</button>
      </div>
      <div className="grid gap-4 md:grid-cols-3 mb-5">
        <KpiCard title="Total Waste" value={money(totalWaste)} helper="This week" />
        <KpiCard title="Product Waste" value={money(productWaste)} helper="Raw ingredients" />
        <KpiCard title="Production Waste" value={money(productionWaste)} helper="Prepared items" />
      </div>
      {showForm && (
        <Card>
          <h3 className="mb-4 font-black text-lg">Log Waste</h3>
          <div className="mb-4 flex gap-2">
            <button onClick={() => handleTypeChange("product")} className={`rounded-xl px-4 py-2 text-sm font-bold transition ${wasteType === "product" ? "bg-[#1a3a5c] text-white" : "bg-[#F7F8FA] text-[#718096] hover:bg-[#E2E8F0]"}`}>Stock Product</button>
            <button onClick={() => handleTypeChange("production")} className={`rounded-xl px-4 py-2 text-sm font-bold transition ${wasteType === "production" ? "bg-[#1a3a5c] text-white" : "bg-[#F7F8FA] text-[#718096] hover:bg-[#E2E8F0]"}`}>Production Item</button>
          </div>
          <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-3">
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">DATE</p><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">{wasteType === "product" ? "PRODUCT" : "PRODUCTION ENTRY"}</p>
              <select value={refId} onChange={(e) => handleRefChange(Number(e.target.value))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none text-sm">
                {currentOptions.length === 0 ? <option>No entries available</option> : currentOptions.map(o => <option key={o.id} value={o.id}>{o.name}</option>)}
              </select>
            </div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">REASON</p>
              <select value={reason} onChange={(e) => setReason(e.target.value)} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none">
                {WASTE_REASONS.map(r => <option key={r}>{r}</option>)}
              </select>
            </div>
            {wasteType === "product" && (
              <>
                <div><p className="mb-1 text-xs font-semibold text-[#718096]">QUANTITY</p><input type="number" value={quantity} onChange={(e) => { setQuantity(Number(e.target.value)); const p = products.find(p => p.id === refId); if (p) setEstimatedCost(p.price * Number(e.target.value)); }} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
                <div><p className="mb-1 text-xs font-semibold text-[#718096]">UNIT</p><select value={unit} onChange={(e) => setUnit(e.target.value)} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none">{UNITS.map(u => <option key={u}>{u}</option>)}</select></div>
              </>
            )}
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">ESTIMATED COST £</p><input type="number" value={estimatedCost} onChange={(e) => setEstimatedCost(Number(e.target.value))} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div className="lg:col-span-3"><p className="mb-1 text-xs font-semibold text-[#718096]">NOTES</p><input value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Optional details" className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
          </div>
          <div className="mt-4 flex gap-3">
            <button onClick={addEntry} className="rounded-xl bg-[#1a3a5c] px-5 py-2 font-bold text-white">Save</button>
            <button onClick={() => setShowForm(false)} className="rounded-xl border border-[#E2E8F0] px-5 py-2 font-bold text-[#718096]">Cancel</button>
          </div>
        </Card>
      )}
      <div className="mt-4 grid gap-5">
        {Object.keys(grouped).length === 0 ? <Card><div className="py-8 text-center text-[#718096]">No waste logged yet.</div></Card> :
          Object.keys(grouped).sort((a, b) => b.localeCompare(a)).map(date => (
            <div key={date}>
              <p className="mb-2 text-sm font-black text-[#718096] uppercase tracking-wide">{date}</p>
              <div className="grid gap-3">
                {grouped[date].map((e) => (
                  <Card key={e.id}>
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <span className={`rounded-full px-2 py-0.5 text-xs font-bold ${e.type === "product" ? "bg-[#E6F1FB] text-[#185FA5]" : "bg-[#FDE8C8] text-[#E07B00]"}`}>
                          {e.type === "product" ? "Stock" : "Production"}
                        </span>
                        <div>
                          <h3 className="font-black">{e.refName}</h3>
                          <p className="text-sm text-[#718096]">{e.reason}{e.notes ? ` · ${e.notes}` : ""}{e.type === "product" ? ` · ${e.quantity} ${e.unit}` : ""}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-4">
                        <div className="text-right"><p className="text-lg font-black text-red-500">{money(e.estimatedCost)}</p><p className="text-xs text-[#718096]">estimated loss</p></div>
                        <button onClick={() => deleteEntry(e.id)} className="rounded-lg bg-red-50 p-2 text-red-400 hover:bg-red-100"><Trash2 size={14} /></button>
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
            </div>
          ))}
      </div>
    </>
  );
}

function Sales({ entries, setEntries }: { entries: SaleEntry[]; setEntries: React.Dispatch<React.SetStateAction<SaleEntry[]>> }) {
  const [mode, setMode] = useState<"import" | "manual" | null>(null);
  const [importMsg, setImportMsg] = useState("");
  const [date, setDate] = useState(new Date().toISOString().split("T")[0]);
  const [source, setSource] = useState("Manual entry");
  const [notes, setNotes] = useState("");
  const [items, setItems] = useState<SaleItem[]>([]);
  const [newItemName, setNewItemName] = useState("");
  const [newItemCat, setNewItemCat] = useState("Mains");
  const [newItemQty, setNewItemQty] = useState(1);
  const [newItemPrice, setNewItemPrice] = useState(0);
  const fileRef = useRef<HTMLInputElement>(null);

  const addItem = () => {
    if (!newItemName.trim()) return;
    setItems(prev => [...prev, { id: Date.now(), name: newItemName, category: newItemCat, quantity: newItemQty, unitPrice: newItemPrice, total: newItemQty * newItemPrice }]);
    setNewItemName(""); setNewItemQty(1); setNewItemPrice(0);
  };

  const saveSale = () => {
    if (items.length === 0) return;
    setEntries(prev => [{ id: Date.now(), date, source, items, totalRevenue: items.reduce((s, i) => s + i.total, 0), notes }, ...prev]);
    setItems([]); setNotes(""); setMode(null);
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportMsg("Reading file...");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const importedItems: SaleItem[] = [];
        let totalRevenue = 0;
        let headerRow = -1;
        for (let i = 0; i < Math.min(rows.length, 10); i++) {
          const row = rows[i].map((c: any) => String(c).toLowerCase());
          if (row.some((c: string) => c.includes("item") || c.includes("product") || c.includes("name"))) { headerRow = i; break; }
        }
        if (headerRow >= 0) {
          const headers = rows[headerRow].map((c: any) => String(c).toLowerCase());
          const nameCol = headers.findIndex((h: string) => h.includes("item") || h.includes("product") || h.includes("name"));
          const qtyCol = headers.findIndex((h: string) => h.includes("qty") || h.includes("quantity"));
          const priceCol = headers.findIndex((h: string) => h.includes("price") || h.includes("unit"));
          const totalCol = headers.findIndex((h: string) => h.includes("total") || h.includes("revenue") || h.includes("amount"));
          for (let i = headerRow + 1; i < rows.length; i++) {
            const row = rows[i];
            const name = String(row[nameCol] || "").trim();
            if (!name) continue;
            const qty = qtyCol >= 0 ? parseFloat(String(row[qtyCol])) || 1 : 1;
            const price = priceCol >= 0 ? parseFloat(String(row[priceCol])) || 0 : 0;
            const total = totalCol >= 0 ? parseFloat(String(row[totalCol])) || qty * price : qty * price;
            importedItems.push({ id: Date.now() + i, name, category: "Imported", quantity: qty, unitPrice: price, total });
            totalRevenue += total;
          }
        }
        if (importedItems.length === 0) { setImportMsg("❌ Could not read items. Try manual entry."); return; }
        setEntries(prev => [{ id: Date.now(), date: new Date().toISOString().split("T")[0], source: file.name, items: importedItems, totalRevenue, notes: "" }, ...prev]);
        setImportMsg(`✅ Imported ${importedItems.length} items. Total: ${money(totalRevenue)}`);
        if (fileRef.current) fileRef.current.value = "";
      } catch { setImportMsg("❌ Error reading file."); }
    };
    reader.readAsArrayBuffer(file);
  };

  const deleteEntry = (id: number) => setEntries(prev => prev.filter(e => e.id !== id));
  const totalSales = entries.reduce((s, e) => s + e.totalRevenue, 0);
  const totalItems = entries.reduce((s, e) => s + e.items.reduce((a, i) => a + i.quantity, 0), 0);
  const grouped = entries.reduce((acc, e) => { if (!acc[e.date]) acc[e.date] = []; acc[e.date].push(e); return acc; }, {} as Record<string, SaleEntry[]>);

  return (
    <>
      <div className="mb-5 flex items-center justify-between">
        <div><h2 className="text-3xl font-black">Sales</h2><p className="text-[#718096]">Log sales from your POS or enter manually.</p></div>
      </div>
      <div className="grid gap-4 md:grid-cols-3 mb-5">
        <KpiCard title="Total Sales" value={money(totalSales)} helper="This week" />
        <KpiCard title="Items Sold" value={totalItems} helper="This week" />
        <KpiCard title="Sale Entries" value={entries.length} helper="Logged this week" />
      </div>
      {!mode && (
        <div className="grid gap-4 md:grid-cols-3 mb-5">
          <label className="flex cursor-pointer flex-col items-center gap-3 rounded-2xl border-2 border-dashed border-[#1a3a5c] p-6 hover:bg-[#1a3a5c]/5 transition">
            <Upload size={28} className="text-[#1a3a5c]" />
            <div className="text-center"><p className="font-black text-[#1a3a5c]">Import from file</p><p className="text-xs text-[#718096]">Excel, CSV — any POS export</p></div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleImport} />
          </label>
          <button onClick={() => setMode("manual")} className="flex flex-col items-center gap-3 rounded-2xl border-2 border-[#E2E8F0] bg-white p-6 hover:border-[#1a3a5c] transition">
            <Plus size={28} className="text-[#1a3a5c]" />
            <div className="text-center"><p className="font-black text-[#1a3a5c]">Manual entry</p><p className="text-xs text-[#718096]">Add items one by one</p></div>
          </button>
          <div className="flex flex-col items-center gap-3 rounded-2xl border-2 border-dashed border-[#E2E8F0] bg-[#F7F8FA] p-6">
            <FileSpreadsheet size={28} className="text-[#718096]" />
            <div className="text-center"><p className="font-black text-[#718096]">AI Reading</p><p className="text-xs text-[#718096]">PDF / photo — coming soon</p></div>
          </div>
        </div>
      )}
      {importMsg && <div className={`mb-4 rounded-xl px-4 py-3 text-sm font-semibold ${importMsg.startsWith("✅") ? "bg-[#D4EDE0] text-[#4A7C59]" : "bg-red-50 text-red-600"}`}>{importMsg}</div>}
      {mode === "manual" && (
        <Card>
          <div className="flex items-center justify-between mb-4">
            <h3 className="font-black text-lg">Manual Sale Entry</h3>
            <button onClick={() => setMode(null)} className="rounded-lg bg-[#F7F8FA] p-2 text-[#718096] hover:bg-[#E2E8F0]"><X size={16} /></button>
          </div>
          <div className="grid gap-3 md:grid-cols-3 mb-4">
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">DATE</p><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">SOURCE</p><input value={source} onChange={(e) => setSource(e.target.value)} placeholder="e.g. Tevalis, Square..." className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">NOTES</p><input value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Optional" className="h-11 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
          </div>
          <div className="mb-3 grid gap-2">
            {items.map(i => (
              <div key={i.id} className="flex items-center justify-between rounded-xl bg-[#F7F8FA] p-3">
                <div><b>{i.name}</b><p className="text-xs text-[#718096]">{i.category} · {i.quantity} × {money(i.unitPrice)}</p></div>
                <div className="flex items-center gap-3"><b className="text-[#1a3a5c]">{money(i.total)}</b><button onClick={() => setItems(prev => prev.filter(x => x.id !== i.id))} className="rounded-lg bg-red-50 p-1.5 text-red-400 hover:bg-red-100"><X size={13} /></button></div>
              </div>
            ))}
          </div>
          <div className="grid gap-2 md:grid-cols-5 items-end border-t border-[#E2E8F0] pt-4">
            <div className="md:col-span-2"><p className="mb-1 text-xs font-semibold text-[#718096]">ITEM NAME</p><input value={newItemName} onChange={(e) => setNewItemName(e.target.value)} placeholder="e.g. Ribeye 300g" className="h-10 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">CATEGORY</p><select value={newItemCat} onChange={(e) => setNewItemCat(e.target.value)} className="h-10 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm">{SALE_CATEGORIES.map(c => <option key={c}>{c}</option>)}</select></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">QTY</p><input type="number" value={newItemQty} onChange={(e) => setNewItemQty(Number(e.target.value))} className="h-10 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm" /></div>
            <div><p className="mb-1 text-xs font-semibold text-[#718096]">UNIT PRICE £</p><input type="number" value={newItemPrice} onChange={(e) => setNewItemPrice(Number(e.target.value))} className="h-10 w-full rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-3 outline-none text-sm" /></div>
          </div>
          <div className="mt-3 flex gap-3">
            <button onClick={addItem} className="rounded-xl border border-[#1a3a5c] px-4 py-2 text-sm font-bold text-[#1a3a5c]"><Plus size={14} className="inline mr-1" />Add item</button>
          </div>
          {items.length > 0 && (
            <div className="mt-4 flex items-center justify-between border-t border-[#E2E8F0] pt-4">
              <p className="font-black text-lg">Total: {money(items.reduce((s, i) => s + i.total, 0))}</p>
              <button onClick={saveSale} className="flex items-center gap-2 rounded-xl bg-[#1a3a5c] px-5 py-2 font-bold text-white"><Save size={16} /> Save Sale</button>
            </div>
          )}
        </Card>
      )}
      <div className="mt-5 grid gap-5">
        {Object.keys(grouped).length === 0 ? <Card><div className="py-8 text-center text-[#718096]">No sales logged yet.</div></Card> :
          Object.keys(grouped).sort((a, b) => b.localeCompare(a)).map(date => (
            <div key={date}>
              <p className="mb-2 text-sm font-black text-[#718096] uppercase tracking-wide">{date}</p>
              <div className="grid gap-3">
                {grouped[date].map(e => (
                  <Card key={e.id}>
                    <div className="flex items-start justify-between">
                      <div>
                        <div className="flex items-center gap-2 mb-1">
                          <h3 className="font-black">{e.source}</h3>
                          <span className="rounded-full bg-[#D4EDE0] px-2 py-0.5 text-xs font-bold text-[#4A7C59]">{e.items.length} items</span>
                        </div>
                        <p className="text-sm text-[#718096]">{e.notes || "No notes"}</p>
                        <div className="mt-2 flex flex-wrap gap-1">
                          {e.items.slice(0, 4).map(i => <span key={i.id} className="rounded-full bg-[#F7F8FA] px-2 py-0.5 text-xs text-[#718096]">{i.name} ×{i.quantity}</span>)}
                          {e.items.length > 4 && <span className="rounded-full bg-[#F7F8FA] px-2 py-0.5 text-xs text-[#718096]">+{e.items.length - 4} more</span>}
                        </div>
                      </div>
                      <div className="flex items-center gap-4">
                        <div className="text-right"><p className="text-2xl font-black text-[#4A7C59]">{money(e.totalRevenue)}</p><p className="text-xs text-[#718096]">total revenue</p></div>
                        <button onClick={() => deleteEntry(e.id)} className="rounded-lg bg-red-50 p-2 text-red-400 hover:bg-red-100"><Trash2 size={14} /></button>
                      </div>
                    </div>
                  </Card>
                ))}
              </div>
            </div>
          ))}
      </div>
    </>
  );
}

function Reports({ products, productionEntries, wasteEntries, salesEntries }: { products: Product[]; productionEntries: ProductionEntry[]; wasteEntries: WasteEntry[]; salesEntries: SaleEntry[] }) {
  const today = new Date().toISOString().split("T")[0];
  const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().split("T")[0];
  const [dateFrom, setDateFrom] = useState(weekAgo);
  const [dateTo, setDateTo] = useState(today);

  const inRange = (date: string) => date >= dateFrom && date <= dateTo;

  const filteredProduction = productionEntries.filter(e => inRange(e.date));
  const filteredWaste = wasteEntries.filter(e => inRange(e.date));
  const filteredSales = salesEntries.filter(e => inRange(e.date));

  const totalSales = filteredSales.reduce((s, e) => s + e.totalRevenue, 0);
  const totalWaste = filteredWaste.reduce((s, e) => s + e.estimatedCost, 0);
  const totalPortions = filteredProduction.reduce((s, e) => s + e.portions, 0);
  const currentStockValue = products.reduce((s, p) => s + p.realStock * p.price, 0);
  const theoreticalStockValue = products.reduce((s, p) => s + p.theoreticalStock * p.price, 0);
  const stockDiff = currentStockValue - theoreticalStockValue;

  const downloadExcel = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Balance
    const balanceData = [
      ["ARKA WEEKLY CLOSE", "", ""],
      ["Period", `${dateFrom} to ${dateTo}`, ""],
      ["", "", ""],
      ["CONCEPT", "VALUE £", "NOTES"],
      ["Current Stock Value (Real)", currentStockValue.toFixed(2), "Real count"],
      ["Theoretical Stock Value", theoreticalStockValue.toFixed(2), "Expected"],
      ["Stock Difference", stockDiff.toFixed(2), stockDiff >= 0 ? "Surplus" : "Deficit"],
      ["Total Sales", totalSales.toFixed(2), `${filteredSales.length} entries`],
      ["Total Waste", totalWaste.toFixed(2), `${filteredWaste.length} entries`],
      ["Total Portions Produced", totalPortions, `${filteredProduction.length} entries`],
      ["Net (Sales - Waste)", (totalSales - totalWaste).toFixed(2), ""],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(balanceData), "Balance");

    // Sheet 2: Production
    const prodData = [["Date", "Recipe", "Portions", "Notes"], ...filteredProduction.map(e => [e.date, e.recipeName, e.portions, e.notes])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(prodData), "Production");

    // Sheet 3: Waste
    const wasteData = [["Date", "Type", "Item", "Qty", "Unit", "Reason", "Cost £", "Notes"], ...filteredWaste.map(e => [e.date, e.type, e.refName, e.quantity, e.unit, e.reason, e.estimatedCost.toFixed(2), e.notes])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wasteData), "Waste");

    // Sheet 4: Sales
    const salesData = [["Date", "Source", "Items", "Total £", "Notes"], ...filteredSales.map(e => [e.date, e.source, e.items.length, e.totalRevenue.toFixed(2), e.notes])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(salesData), "Sales");

    // Sheet 5: Stock
    const stockData = [["Product", "Category", "Unit", "Price £", "Real Stock", "Theoretical Stock", "PAR", "Stock Value £"], ...products.map(p => [p.name, p.category, p.unit, p.price.toFixed(2), p.realStock, p.theoreticalStock, p.parLevel, (p.realStock * p.price).toFixed(2)])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(stockData), "Stock");

    XLSX.writeFile(wb, `ARKA_WeeklyClose_${dateFrom}_${dateTo}.xlsx`);
  };

  return (
    <>
      <div className="mb-5 flex items-center justify-between">
        <div><h2 className="text-3xl font-black">Reports</h2><p className="text-[#718096]">Weekly close — balance of stock, production, waste and sales.</p></div>
        <button onClick={downloadExcel} className="flex items-center gap-2 rounded-xl bg-[#4A7C59] px-5 py-3 font-bold text-white"><Download size={18} /> Download Excel</button>
      </div>

      {/* Date selector */}
      <Card>
        <div className="flex flex-wrap items-center gap-4">
          <div><p className="mb-1 text-xs font-semibold text-[#718096]">FROM</p><input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} className="h-11 rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
          <div><p className="mb-1 text-xs font-semibold text-[#718096]">TO</p><input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} className="h-11 rounded-xl border border-[#E2E8F0] bg-[#F7F8FA] px-4 outline-none" /></div>
          <div className="flex gap-2 mt-4">
            <button onClick={() => { const d = new Date(); setDateTo(d.toISOString().split("T")[0]); d.setDate(d.getDate() - 7); setDateFrom(d.toISOString().split("T")[0]); }} className="rounded-xl bg-[#F7F8FA] px-4 py-2 text-sm font-bold text-[#718096] hover:bg-[#E2E8F0]">Last 7 days</button>
            <button onClick={() => { const d = new Date(); setDateTo(d.toISOString().split("T")[0]); d.setDate(d.getDate() - 30); setDateFrom(d.toISOString().split("T")[0]); }} className="rounded-xl bg-[#F7F8FA] px-4 py-2 text-sm font-bold text-[#718096] hover:bg-[#E2E8F0]">Last 30 days</button>
          </div>
        </div>
      </Card>

      {/* Balance */}
      <div className="mt-5 grid gap-4 md:grid-cols-2 xl:grid-cols-4">
        <KpiCard title="Total Sales" value={money(totalSales)} helper={`${filteredSales.length} entries`} />
        <KpiCard title="Total Waste" value={money(totalWaste)} helper={`${filteredWaste.length} entries`} />
        <KpiCard title="Portions Produced" value={totalPortions} helper={`${filteredProduction.length} entries`} />
        <KpiCard title="Net (Sales − Waste)" value={money(totalSales - totalWaste)} helper="This period" />
      </div>

      <div className="mt-5 grid gap-4 xl:grid-cols-2">
        <Card>
          <h3 className="mb-4 text-xl font-black">Stock Balance</h3>
          <div className="grid gap-3">
            <div className="flex justify-between rounded-xl bg-[#F7F8FA] p-3"><span className="font-semibold">Real Stock Value</span><span className="font-black">{money(currentStockValue)}</span></div>
            <div className="flex justify-between rounded-xl bg-[#F7F8FA] p-3"><span className="font-semibold">Theoretical Stock Value</span><span className="font-black">{money(theoreticalStockValue)}</span></div>
            <div className={`flex justify-between rounded-xl p-3 ${stockDiff >= 0 ? "bg-[#D4EDE0]" : "bg-red-50"}`}>
              <span className="font-black">Difference</span>
              <span className={`font-black ${stockDiff >= 0 ? "text-[#4A7C59]" : "text-red-500"}`}>{money(stockDiff)}</span>
            </div>
          </div>
        </Card>

        <Card>
          <h3 className="mb-4 text-xl font-black">Weekly Summary</h3>
          <div className="grid gap-3">
            <div className="flex justify-between rounded-xl bg-[#F7F8FA] p-3"><span className="font-semibold">Sales Revenue</span><span className="font-black text-[#4A7C59]">+{money(totalSales)}</span></div>
            <div className="flex justify-between rounded-xl bg-[#F7F8FA] p-3"><span className="font-semibold">Waste Cost</span><span className="font-black text-red-500">−{money(totalWaste)}</span></div>
            <div className="flex justify-between rounded-xl bg-[#1a3a5c] p-3 text-white"><span className="font-black">Net Result</span><span className="font-black text-xl">{money(totalSales - totalWaste)}</span></div>
          </div>
        </Card>
      </div>

      {/* Production breakdown */}
      {filteredProduction.length > 0 && (
        <div className="mt-5">
          <Card>
            <h3 className="mb-4 text-xl font-black">Production in Period</h3>
            <div className="grid gap-2">
              {filteredProduction.map(e => (
                <div key={e.id} className="flex items-center justify-between rounded-xl bg-[#F7F8FA] p-3">
                  <div><b>{e.recipeName}</b><p className="text-xs text-[#718096]">{e.date} · {e.notes || "No notes"}</p></div>
                  <span className="font-black text-[#1a3a5c]">{e.portions} portions</span>
                </div>
              ))}
            </div>
          </Card>
        </div>
      )}

      {filteredProduction.length === 0 && filteredWaste.length === 0 && filteredSales.length === 0 && (
        <div className="mt-5">
          <Card><div className="py-8 text-center text-[#718096]">No data for the selected period. Log production, waste and sales first.</div></Card>
        </div>
      )}
    </>
  );
}

export default function ArkaApp() {
  const [active, setActive] = useState("dashboard");
  const [products, setProducts] = useState<Product[]>(() => { const s = localStorage.getItem("arka_products"); return s ? JSON.parse(s) : initialProducts; });
  const [recipes, setRecipes] = useState<Recipe[]>(() => { const s = localStorage.getItem("arka_recipes"); return s ? JSON.parse(s) : initialRecipes; });
  const [productionEntries, setProductionEntries] = useState<ProductionEntry[]>(() => { const s = localStorage.getItem("arka_production"); return s ? JSON.parse(s) : []; });
  const [wasteEntries, setWasteEntries] = useState<WasteEntry[]>(() => { const s = localStorage.getItem("arka_waste"); return s ? JSON.parse(s) : []; });
  const [salesEntries, setSalesEntries] = useState<SaleEntry[]>(() => { const s = localStorage.getItem("arka_sales"); return s ? JSON.parse(s) : []; });

  React.useEffect(() => { localStorage.setItem("arka_products", JSON.stringify(products)); }, [products]);
  React.useEffect(() => { localStorage.setItem("arka_recipes", JSON.stringify(recipes)); }, [recipes]);
  React.useEffect(() => { localStorage.setItem("arka_production", JSON.stringify(productionEntries)); }, [productionEntries]);
  React.useEffect(() => { localStorage.setItem("arka_waste", JSON.stringify(wasteEntries)); }, [wasteEntries]);
  React.useEffect(() => { localStorage.setItem("arka_sales", JSON.stringify(salesEntries)); }, [salesEntries]);

  return (
    <Shell active={active} setActive={setActive}>
      {active === "dashboard" && <Dashboard products={products} productionEntries={productionEntries} wasteEntries={wasteEntries} salesEntries={salesEntries} />}
      {active === "stock" && <Stocktake products={products} setProducts={setProducts} />}
      {active === "recipes" && <Recipes products={products} recipes={recipes} setRecipes={setRecipes} />}
      {active === "production" && <Production entries={productionEntries} setEntries={setProductionEntries} recipes={recipes} />}
      {active === "waste" && <Waste products={products} entries={wasteEntries} setEntries={setWasteEntries} productionEntries={productionEntries} />}
      {active === "sales" && <Sales entries={salesEntries} setEntries={setSalesEntries} />}
      {active === "reports" && <Reports products={products} productionEntries={productionEntries} wasteEntries={wasteEntries} salesEntries={salesEntries} />}
      {active === "users" && <div className="p-5"><h2 className="text-3xl font-black">Users</h2><p className="mt-2 text-[#718096]">Coming soon.</p></div>}
      {active === "settings" && <div className="p-5"><h2 className="text-3xl font-black">Settings</h2><p className="mt-2 text-[#718096]">Coming soon.</p></div>}
    </Shell>
  );
}