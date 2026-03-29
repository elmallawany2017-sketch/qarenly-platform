import React, { useMemo, useRef, useState } from "react";
import {
  Upload,
  Search,
  BarChart3,
  FileSpreadsheet,
  ArrowUpDown,
  CheckCircle2,
  XCircle,
  Equal,
  Store,
  ShieldCheck,
  Users,
  Trash2,
  LogIn,
  UserPlus,
  LayoutDashboard,
  Crown,
  Download,
  TrendingUp,
  Briefcase,
  Bell,
} from "lucide-react";
import * as XLSX from "xlsx";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { motion } from "framer-motion";

const currency = (n: number | string) => {
  const num = Number(n || 0);
  return new Intl.NumberFormat("en-EG", { maximumFractionDigits: 2 }).format(num);
};

const normalizeName = (value: string) =>
  String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[\u0640]/g, "")
    .replace(/[أإآ]/g, "ا")
    .replace(/[ة]/g, "ه")
    .replace(/[ى]/g, "ي")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();

const calcFinalPrice = (price: number, discount: number) => {
  const p = Number(price || 0);
  const d = Number(discount || 0);
  return +(p - p * (d / 100)).toFixed(2);
};

const levenshtein = (a: string, b: string) => {
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i += 1) dp[i][0] = i;
  for (let j = 0; j <= n; j += 1) dp[0][j] = j;
  for (let i = 1; i <= m; i += 1) {
    for (let j = 1; j <= n; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[m][n];
};

const similarityScore = (a: string, b: string) => {
  const x = normalizeName(a);
  const y = normalizeName(b);
  if (!x || !y) return 0;
  if (x === y) return 1;
  if (x.includes(y) || y.includes(x)) return Math.min(x.length, y.length) / Math.max(x.length, y.length);
  const distance = levenshtein(x, y);
  const maxLen = Math.max(x.length, y.length);
  return maxLen === 0 ? 1 : 1 - distance / maxLen;
};

const bestMatch = (target: string, candidates: string[], threshold = 0.8) => {
  let bestKey: string | null = null;
  let best = 0;
  for (const candidate of candidates) {
    const score = similarityScore(target, candidate);
    if (score > best) {
      best = score;
      bestKey = candidate;
    }
  }
  if (!bestKey || best < threshold) return null;
  return { key: bestKey, score: best };
};

const guessColumn = (headers: string[], kind: "name" | "price" | "discount") => {
  const normalized = headers.map((h) => ({ raw: h, key: normalizeName(h) }));
  const patterns = {
    name: ["name", "product name", "اسم الصنف", "الصنف", "اسم المنتج"],
    price: ["price", "السعر", "سعر", "unit price", "buy price"],
    discount: ["discount", "الخصم", "discount %", "خصم", "نسبه الخصم"],
  }[kind];

  for (const pattern of patterns) {
    const found = normalized.find((h) => h.key === normalizeName(pattern) || h.key.includes(normalizeName(pattern)));
    if (found) return found.raw;
  }
  return headers[0] || "";
};

const readExcelRaw = async (file: File) => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" }) as Record<string, unknown>[];
  const headers = rows.length ? Object.keys(rows[0]) : [];
  return { rows, headers, fileName: file.name };
};

const mapRowsFromSelection = (
  rows: Record<string, unknown>[],
  mapping: { name: string; price: string; discount: string }
) => {
  return rows.map((row, index) => {
    const productName = String(row[mapping.name] ?? "").trim();
    const price = Number(row[mapping.price]);
    const discount = Number(row[mapping.discount]);
    const errors: string[] = [];

    if (!productName) errors.push("Missing product name");
    if (Number.isNaN(price) || price <= 0) errors.push("Invalid price");
    if (Number.isNaN(discount) || discount < 0 || discount > 100) errors.push("Invalid discount");

    return {
      rowNumber: index + 2,
      productName,
      normalizedName: normalizeName(productName),
      price,
      discount,
      finalPrice: calcFinalPrice(price, discount),
      errors,
      valid: errors.length === 0,
    };
  });
};

const downloadCsv = (rows: Record<string, unknown>[], fileName: string) => {
  const header = Object.keys(rows[0] || {});
  const csv = [header.join(",")]
    .concat(
      rows.map((row) =>
        header
          .map((key) => {
            const value = String(row[key] ?? "").replace(/"/g, '""');
            return `"${value}"`;
          })
          .join(",")
      )
    )
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
};

const seedCompanies = [
  {
    id: crypto.randomUUID(),
    companyName: "Alpha Pharma",
    email: "alpha@qarenly.com",
    plan: "Pro",
    offers: [
      { productName: "Panadol Extra", normalizedName: normalizeName("Panadol Extra"), price: 100, discount: 20, finalPrice: 80 },
      { productName: "Augmentin 1g", normalizedName: normalizeName("Augmentin 1g"), price: 180, discount: 12, finalPrice: 158.4 },
      { productName: "Cetal 500", normalizedName: normalizeName("Cetal 500"), price: 30, discount: 5, finalPrice: 28.5 },
    ],
  },
  {
    id: crypto.randomUUID(),
    companyName: "Trust Med",
    email: "trust@qarenly.com",
    plan: "Business",
    offers: [
      { productName: "Panadol Exra", normalizedName: normalizeName("Panadol Exra"), price: 98, discount: 15, finalPrice: 83.3 },
      { productName: "Augmentin 1 gm", normalizedName: normalizeName("Augmentin 1 gm"), price: 175, discount: 8, finalPrice: 161 },
      { productName: "Brufen 400", normalizedName: normalizeName("Brufen 400"), price: 48, discount: 10, finalPrice: 43.2 },
    ],
  },
  {
    id: crypto.randomUUID(),
    companyName: "Market Plus",
    email: "market@qarenly.com",
    plan: "Free",
    offers: [
      { productName: "Panadol Extra", normalizedName: normalizeName("Panadol Extra"), price: 101, discount: 22, finalPrice: 78.78 },
      { productName: "Cetal500", normalizedName: normalizeName("Cetal500"), price: 31, discount: 8, finalPrice: 28.52 },
      { productName: "Brufen 400", normalizedName: normalizeName("Brufen 400"), price: 47, discount: 5, finalPrice: 44.65 },
    ],
  },
];

const seedUsers = [
  { id: "admin-1", name: "Admin", email: "admin@qarenly.com", password: "123456", role: "admin", companyName: "Qarenly" },
  { id: "supplier-1", name: "Alpha Pharma", email: "alpha@qarenly.com", password: "123456", role: "supplier", companyName: "Alpha Pharma" },
  { id: "supplier-2", name: "Trust Med", email: "trust@qarenly.com", password: "123456", role: "supplier", companyName: "Trust Med" },
];

const StatCard = ({ title, value, icon: Icon, hint }: any) => (
  <Card className="rounded-2xl shadow-sm">
    <CardContent className="p-5">
      <div className="flex items-start justify-between gap-3">
        <div>
          <p className="text-sm text-muted-foreground">{title}</p>
          <p className="mt-2 break-words text-2xl font-bold">{value}</p>
          {hint ? <p className="mt-1 text-xs text-muted-foreground">{hint}</p> : null}
        </div>
        <div className="rounded-2xl border bg-white p-3">
          <Icon className="h-5 w-5" />
        </div>
      </div>
    </CardContent>
  </Card>
);

const SectionTitle = ({ title, desc, action }: any) => (
  <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
    <div>
      <h2 className="text-2xl font-bold">{title}</h2>
      {desc ? <p className="mt-1 text-sm text-muted-foreground">{desc}</p> : null}
    </div>
    {action}
  </div>
);

export default function QarenlyPlatform() {
  const [companies, setCompanies] = useState(seedCompanies);
  const [users, setUsers] = useState(seedUsers);
  const [authMode, setAuthMode] = useState("login");
  const [currentUser, setCurrentUser] = useState<any>(null);
  const [loginForm, setLoginForm] = useState({ email: "admin@qarenly.com", password: "123456" });
  const [registerForm, setRegisterForm] = useState({ name: "", email: "", password: "", companyName: "" });
  const [authError, setAuthError] = useState("");
  const [activeTab, setActiveTab] = useState("market");
  const [companyName, setCompanyName] = useState("");
  const [uploadReport, setUploadReport] = useState<any>(null);
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState("discount");
  const [filterCompany, setFilterCompany] = useState("all");
  const [compareA, setCompareA] = useState<any>(null);
  const [compareB, setCompareB] = useState<any>(null);
  const [compareResult, setCompareResult] = useState<any>(null);
  const [uploadDraft, setUploadDraft] = useState<any>(null);
  const [compareDraftA, setCompareDraftA] = useState<any>(null);
  const [compareDraftB, setCompareDraftB] = useState<any>(null);

  const companyFileRef = useRef<HTMLInputElement>(null);
  const compareARef = useRef<HTMLInputElement>(null);
  const compareBRef = useRef<HTMLInputElement>(null);

  const allOffers = useMemo(() => {
    return companies.flatMap((company: any) =>
      company.offers.map((offer: any) => ({ ...offer, companyName: company.companyName, companyId: company.id, plan: company.plan }))
    );
  }, [companies]);

  const groupedResults = useMemo(() => {
    const query = normalizeName(search);
    let data = allOffers;
    if (query) {
      data = data.filter((offer: any) => offer.normalizedName.includes(query) || normalizeName(offer.companyName).includes(query));
    }
    if (filterCompany !== "all") {
      data = data.filter((offer: any) => offer.companyId === filterCompany);
    }

    const groups: any[] = [];
    for (const row of data) {
      let group = groups.find((g) => similarityScore(g.anchor, row.productName) >= 0.8);
      if (!group) {
        group = { key: crypto.randomUUID(), anchor: row.productName, displayName: row.productName, rows: [] };
        groups.push(group);
      }
      group.rows.push(row);
    }

    groups.forEach((group) => {
      group.rows.sort((a: any, b: any) => {
        if (sortBy === "discount") return b.discount - a.discount || a.finalPrice - b.finalPrice;
        if (sortBy === "final") return a.finalPrice - b.finalPrice || b.discount - a.discount;
        return a.productName.localeCompare(b.productName);
      });
    });

    groups.sort((a, b) => a.displayName.localeCompare(b.displayName));
    return groups;
  }, [allOffers, search, sortBy, filterCompany]);

  const productCount = groupedResults.length;
  const avgDiscount = allOffers.length ? (allOffers.reduce((sum: number, row: any) => sum + row.discount, 0) / allOffers.length).toFixed(1) : 0;

  const myCompany = useMemo(() => {
    if (!currentUser || currentUser.role !== "supplier") return null;
    return companies.find((c: any) => normalizeName(c.companyName) === normalizeName(currentUser.companyName));
  }, [companies, currentUser]);

  const marketLeaderboard = useMemo(() => {
    return companies
      .map((company: any) => {
        const offers = company.offers || [];
        const avg = offers.length ? offers.reduce((sum: number, row: any) => sum + row.discount, 0) / offers.length : 0;
        return { ...company, avgDiscount: avg.toFixed(1), items: offers.length };
      })
      .sort((a: any, b: any) => Number(b.avgDiscount) - Number(a.avgDiscount));
  }, [companies]);

  const login = () => {
    const found = users.find((u: any) => u.email === loginForm.email && u.password === loginForm.password);
    if (!found) {
      setAuthError("بيانات الدخول غير صحيحة");
      return;
    }
    setAuthError("");
    setCurrentUser(found);
    setActiveTab(found.role === "admin" ? "admin" : "dashboard");
  };

  const register = () => {
    if (!registerForm.name || !registerForm.email || !registerForm.password || !registerForm.companyName) {
      setAuthError("اكمل كل البيانات أولًا");
      return;
    }
    if (users.some((u: any) => u.email === registerForm.email)) {
      setAuthError("هذا البريد مستخدم بالفعل");
      return;
    }
    const newUser = {
      id: crypto.randomUUID(),
      name: registerForm.name,
      email: registerForm.email,
      password: registerForm.password,
      role: "supplier",
      companyName: registerForm.companyName,
    };
    setUsers((prev: any) => [...prev, newUser]);
    setCompanies((prev: any) => [...prev, { id: crypto.randomUUID(), companyName: registerForm.companyName, email: registerForm.email, plan: "Free", offers: [] }]);
    setCurrentUser(newUser);
    setAuthError("");
    setActiveTab("dashboard");
    setRegisterForm({ name: "", email: "", password: "", companyName: "" });
  };

  const logout = () => {
    setCurrentUser(null);
    setActiveTab("market");
  };

  const handleCompanyDraft = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const draft = await readExcelRaw(file);
    const mapping = {
      name: guessColumn(draft.headers, "name"),
      price: guessColumn(draft.headers, "price"),
      discount: guessColumn(draft.headers, "discount"),
    };
    setUploadDraft({ ...draft, mapping });
    event.target.value = "";
  };

  const finalizeUpload = () => {
    if (!uploadDraft) return;
    const rows = mapRowsFromSelection(uploadDraft.rows, uploadDraft.mapping);
    const validRows = rows.filter((r) => r.valid);
    const invalidRows = rows.filter((r) => !r.valid);
    const resolvedCompanyName = currentUser?.role === "supplier" ? currentUser.companyName : companyName.trim() || uploadDraft.fileName.replace(/\.[^.]+$/, "");
    const mappedOffers = validRows.map(({ rowNumber, errors, valid, ...rest }) => rest);

    setCompanies((prev: any) => {
      const existingIndex = prev.findIndex((c: any) => normalizeName(c.companyName) === normalizeName(resolvedCompanyName));
      if (existingIndex >= 0) {
        const clone = [...prev];
        clone[existingIndex] = { ...clone[existingIndex], offers: mappedOffers };
        return clone;
      }
      return [...prev, { id: crypto.randomUUID(), companyName: resolvedCompanyName, email: currentUser?.email || "", plan: "Free", offers: mappedOffers }];
    });

    setUploadReport({ fileName: uploadDraft.fileName, companyName: resolvedCompanyName, total: rows.length, valid: validRows.length, invalid: invalidRows, mapping: uploadDraft.mapping });
    setUploadDraft(null);
    setCompanyName("");
  };

  const handleCompareDraft = async (event: React.ChangeEvent<HTMLInputElement>, which: "A" | "B") => {
    const file = event.target.files?.[0];
    if (!file) return;
    const draft = await readExcelRaw(file);
    const mapping = {
      name: guessColumn(draft.headers, "name"),
      price: guessColumn(draft.headers, "price"),
      discount: guessColumn(draft.headers, "discount"),
    };
    const payload = { ...draft, mapping, label: which === "A" ? "ملف 1" : "ملف 2" };
    if (which === "A") setCompareDraftA(payload);
    else setCompareDraftB(payload);
    event.target.value = "";
  };

  const finalizeCompareFile = (which: "A" | "B") => {
    const draft = which === "A" ? compareDraftA : compareDraftB;
    if (!draft) return;
    const rows = mapRowsFromSelection(draft.rows, draft.mapping).filter((r) => r.valid).map(({ rowNumber, errors, valid, ...rest }) => rest);
    const payload = { name: draft.label, rawName: draft.fileName, rows, mapping: draft.mapping };
    if (which === "A") {
      setCompareA(payload);
      setCompareDraftA(null);
    } else {
      setCompareB(payload);
      setCompareDraftB(null);
    }
  };

  const runComparison = () => {
    if (!compareA || !compareB) return;
    const bKeys = compareB.rows.map((r: any) => r.productName);
    const usedB = new Set<string>();
    const shared: any[] = [];
    const onlyA: any[] = [];

    for (const a of compareA.rows) {
      const match = bestMatch(a.productName, bKeys.filter((name: string) => !usedB.has(name)), 0.8);
      if (!match) {
        onlyA.push(a);
        continue;
      }
      const b = compareB.rows.find((row: any) => row.productName === match.key);
      if (!b) {
        onlyA.push(a);
        continue;
      }
      usedB.add(b.productName);
      let winner = "Equal";
      if (a.discount > b.discount) winner = compareA.name;
      else if (b.discount > a.discount) winner = compareB.name;
      else if (a.finalPrice < b.finalPrice) winner = compareA.name;
      else if (b.finalPrice < a.finalPrice) winner = compareB.name;
      shared.push({
        key: `${a.productName}-${b.productName}`,
        productName: a.productName,
        a,
        b,
        similarity: +(match.score * 100).toFixed(0),
        winner,
      });
    }

    const onlyB = compareB.rows.filter((row: any) => !usedB.has(row.productName));
    const aBetter = shared.filter((r) => r.winner === compareA.name).length;
    const bBetter = shared.filter((r) => r.winner === compareB.name).length;
    const equal = shared.filter((r) => r.winner === "Equal").length;

    setCompareResult({
      shared,
      onlyA,
      onlyB,
      aBetter,
      bBetter,
      equal,
      totalA: compareA.rows.length,
      totalB: compareB.rows.length,
      sharedCount: shared.length,
      threshold: 80,
    });
  };

  const removeCompany = (id: string) => setCompanies((prev: any) => prev.filter((c: any) => c.id !== id));

  const exportComparison = () => {
    if (!compareResult) return;
    const rows = compareResult.shared.map((row: any) => ({
      product_name_file_1: row.a.productName,
      product_name_file_2: row.b.productName,
      similarity_percent: row.similarity,
      file_1_price: row.a.price,
      file_1_discount: row.a.discount,
      file_1_final: row.a.finalPrice,
      file_2_price: row.b.price,
      file_2_discount: row.b.discount,
      file_2_final: row.b.finalPrice,
      winner: row.winner,
    }));
    if (rows.length) downloadCsv(rows, "comparison-results.csv");
  };

  const renderAuth = () => (
    <div className="grid gap-6 lg:grid-cols-[1.1fr_.9fr]">
      <Card className="rounded-3xl border-0 bg-gradient-to-br from-white to-slate-100 shadow-lg">
        <CardContent className="p-8 md:p-10">
          <Badge className="mb-4 rounded-full px-4 py-1 text-sm">Qarenly SaaS</Badge>
          <h1 className="text-3xl font-bold tracking-tight md:text-5xl">منصة ذكية لمقارنة الخصومات والأسعار</h1>
          <p className="mt-4 max-w-2xl text-base text-muted-foreground md:text-lg">
            ارفع Excel وحدد بنفسك عمود الاسم وعمود السعر وعمود الخصم، ثم قارن الأصناف بنسبة تشابه 80% لمراعاة اختلاف الكتابة.
          </p>
          <div className="mt-6 grid gap-4 sm:grid-cols-3">
            <StatCard title="الشركات" value={companies.length} icon={Store} hint="موردون وعملاء" />
            <StatCard title="العروض" value={allOffers.length} icon={BarChart3} hint="إجمالي العروض" />
            <StatCard title="متوسط الخصم" value={`${avgDiscount}%`} icon={TrendingUp} hint="على مستوى السوق" />
          </div>
        </CardContent>
      </Card>

      <Card className="rounded-3xl shadow-lg">
        <CardHeader>
          <CardTitle className="text-2xl">{authMode === "login" ? "تسجيل الدخول" : "إنشاء حساب مورد"}</CardTitle>
          <CardDescription>جرب بحساب الأدمن أو أنشئ حساب مورد جديد.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          {authMode === "login" ? (
            <>
              <Input className="rounded-2xl" placeholder="البريد الإلكتروني" value={loginForm.email} onChange={(e) => setLoginForm((p) => ({ ...p, email: e.target.value }))} />
              <Input className="rounded-2xl" type="password" placeholder="كلمة المرور" value={loginForm.password} onChange={(e) => setLoginForm((p) => ({ ...p, password: e.target.value }))} />
              <Button className="w-full rounded-2xl" onClick={login}><LogIn className="mr-2 h-4 w-4" />دخول</Button>
              <div className="rounded-2xl border p-4 text-sm text-muted-foreground">
                <p>حساب الأدمن: admin@qarenly.com</p>
                <p>كلمة المرور: 123456</p>
              </div>
            </>
          ) : (
            <>
              <Input className="rounded-2xl" placeholder="الاسم" value={registerForm.name} onChange={(e) => setRegisterForm((p) => ({ ...p, name: e.target.value }))} />
              <Input className="rounded-2xl" placeholder="اسم الشركة" value={registerForm.companyName} onChange={(e) => setRegisterForm((p) => ({ ...p, companyName: e.target.value }))} />
              <Input className="rounded-2xl" placeholder="البريد الإلكتروني" value={registerForm.email} onChange={(e) => setRegisterForm((p) => ({ ...p, email: e.target.value }))} />
              <Input className="rounded-2xl" type="password" placeholder="كلمة المرور" value={registerForm.password} onChange={(e) => setRegisterForm((p) => ({ ...p, password: e.target.value }))} />
              <Button className="w-full rounded-2xl" onClick={register}><UserPlus className="mr-2 h-4 w-4" />إنشاء الحساب</Button>
            </>
          )}
          {authError ? <Alert className="rounded-2xl"><AlertTitle>تنبيه</AlertTitle><AlertDescription>{authError}</AlertDescription></Alert> : null}
          <Button variant="ghost" className="w-full rounded-2xl" onClick={() => { setAuthMode(authMode === "login" ? "register" : "login"); setAuthError(""); }}>
            {authMode === "login" ? "إنشاء حساب جديد" : "لديك حساب بالفعل؟ سجل دخول"}
          </Button>
        </CardContent>
      </Card>
    </div>
  );

  const renderTopBar = () => (
    <Card className="rounded-3xl border-0 shadow-lg">
      <CardContent className="flex flex-col gap-5 p-6 md:flex-row md:items-center md:justify-between">
        <div>
          <div className="flex items-center gap-2">
            <Badge className="rounded-full px-4 py-1">{currentUser?.role === "admin" ? "Admin" : "Supplier"}</Badge>
            <Badge variant="secondary" className="rounded-full px-4 py-1">{currentUser?.companyName}</Badge>
          </div>
          <h1 className="mt-3 text-3xl font-bold">مرحبًا، {currentUser?.name}</h1>
          <p className="mt-1 text-sm text-muted-foreground">رفع مرن للأعمدة + مقارنة بنسبة تشابه 80% بين الأصناف.</p>
        </div>
        <div className="flex flex-wrap gap-3">
          <Button variant="outline" className="rounded-2xl" onClick={logout}>تسجيل خروج</Button>
        </div>
      </CardContent>
    </Card>
  );

  const renderMarket = () => (
    <div className="space-y-6">
      <Card className="rounded-3xl shadow-sm">
        <CardHeader>
          <CardTitle className="text-2xl">محرك مقارنة السوق</CardTitle>
          <CardDescription>عند البحث عن أي صنف ستظهر كل الشركات أو العملاء العارضين له، مع اسم كل عميل وبجواره الخصم والسعر والسعر النهائي، مرتبين لسهولة معرفة أعلى خصم.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid gap-4 lg:grid-cols-[1fr_220px_260px]">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground" />
              <Input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="ابحث باسم الصنف أو الشركة..." className="rounded-2xl pl-10" />
            </div>
            <Select value={sortBy} onValueChange={setSortBy}>
              <SelectTrigger className="rounded-2xl"><SelectValue placeholder="ترتيب النتائج" /></SelectTrigger>
              <SelectContent>
                <SelectItem value="discount">أعلى خصم</SelectItem>
                <SelectItem value="final">أقل سعر نهائي</SelectItem>
                <SelectItem value="name">اسم الصنف</SelectItem>
              </SelectContent>
            </Select>
            <Select value={filterCompany} onValueChange={setFilterCompany}>
              <SelectTrigger className="rounded-2xl"><SelectValue placeholder="فلترة بالشركة" /></SelectTrigger>
              <SelectContent>
                <SelectItem value="all">كل الشركات</SelectItem>
                {companies.map((company: any) => (
                  <SelectItem key={company.id} value={company.id}>{company.companyName}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </CardContent>
      </Card>

      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
        <StatCard title="إجمالي الشركات" value={companies.length} icon={Store} />
        <StatCard title="إجمالي العروض" value={allOffers.length} icon={BarChart3} />
        <StatCard title="مجموعات الأصناف" value={productCount} icon={ShieldCheck} hint="بالتجميع الذكي 80%" />
        <StatCard title="متوسط الخصم" value={`${avgDiscount}%`} icon={TrendingUp} />
      </div>

      <div className="space-y-5">
        {groupedResults.length === 0 ? (
          <Card className="rounded-3xl"><CardContent className="p-8 text-center text-muted-foreground">لا توجد نتائج مطابقة.</CardContent></Card>
        ) : (
          groupedResults.map((group) => (
            <Card key={group.key} className="rounded-3xl shadow-sm">
              <CardHeader>
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div>
                    <CardTitle className="text-xl">{group.displayName}</CardTitle>
                    <CardDescription>{group.rows.length} عروض متقاربة</CardDescription>
                  </div>
                  <Badge variant="secondary" className="rounded-full px-4 py-1">
                    أفضل عرض: {sortBy === "final" ? `${currency(group.rows[0].finalPrice)} ج` : `${group.rows[0].discount}%`}
                  </Badge>
                </div>
              </CardHeader>
              <CardContent>
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>الترتيب</TableHead>
                      <TableHead>الشركة / العميل</TableHead>
                      <TableHead>اسم الصنف</TableHead>
                      <TableHead>الخصم</TableHead>
                      <TableHead>السعر</TableHead>
                      <TableHead>السعر النهائي</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {group.rows.map((row: any, index: number) => (
                      <TableRow key={`${row.companyId}-${index}`}>
                        <TableCell>{index + 1}</TableCell>
                        <TableCell className="font-medium">{row.companyName}</TableCell>
                        <TableCell>{row.productName}</TableCell>
                        <TableCell>{row.discount}%</TableCell>
                        <TableCell>{currency(row.price)} ج</TableCell>
                        <TableCell>{currency(row.finalPrice)} ج</TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </CardContent>
            </Card>
          ))
        )}
      </div>
    </div>
  );

  const renderUpload = () => (
    <Card className="rounded-3xl shadow-sm">
      <CardHeader>
        <CardTitle className="text-2xl">رفع ملف شركة</CardTitle>
        <CardDescription>بعد رفع الملف ستختار بنفسك: عمود الاسم + عمود السعر + عمود الخصم.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-5">
        {currentUser?.role !== "supplier" ? (
          <div className="grid gap-4 md:grid-cols-[1fr_auto]">
            <Input value={companyName} onChange={(e) => setCompanyName(e.target.value)} placeholder="اسم الشركة أو العميل" className="rounded-2xl" />
            <Button className="rounded-2xl" onClick={() => companyFileRef.current?.click()}><Upload className="mr-2 h-4 w-4" /> اختر ملف Excel</Button>
          </div>
        ) : (
          <div className="rounded-2xl border p-4 text-sm text-muted-foreground">
            سيتم ربط الملف تلقائيًا بشركتك: <strong>{currentUser.companyName}</strong>
          </div>
        )}

        {currentUser?.role === "supplier" ? (
          <Button className="rounded-2xl" onClick={() => companyFileRef.current?.click()}><Upload className="mr-2 h-4 w-4" /> رفع ملف Excel</Button>
        ) : null}

        <input ref={companyFileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleCompanyDraft} />

        {uploadDraft ? (
          <Card className="rounded-2xl border">
            <CardContent className="grid gap-4 p-5 md:grid-cols-3">
              <div>
                <p className="mb-2 text-sm font-medium">اختار عمود الاسم</p>
                <Select value={uploadDraft.mapping.name} onValueChange={(value) => setUploadDraft((prev: any) => ({ ...prev, mapping: { ...prev.mapping, name: value } }))}>
                  <SelectTrigger className="rounded-2xl"><SelectValue /></SelectTrigger>
                  <SelectContent>{uploadDraft.headers.map((h: string) => <SelectItem key={h} value={h}>{h}</SelectItem>)}</SelectContent>
                </Select>
              </div>
              <div>
                <p className="mb-2 text-sm font-medium">اختار عمود السعر</p>
                <Select value={uploadDraft.mapping.price} onValueChange={(value) => setUploadDraft((prev: any) => ({ ...prev, mapping: { ...prev.mapping, price: value } }))}>
                  <SelectTrigger className="rounded-2xl"><SelectValue /></SelectTrigger>
                  <SelectContent>{uploadDraft.headers.map((h: string) => <SelectItem key={h} value={h}>{h}</SelectItem>)}</SelectContent>
                </Select>
              </div>
              <div>
                <p className="mb-2 text-sm font-medium">اختار عمود الخصم</p>
                <Select value={uploadDraft.mapping.discount} onValueChange={(value) => setUploadDraft((prev: any) => ({ ...prev, mapping: { ...prev.mapping, discount: value } }))}>
                  <SelectTrigger className="rounded-2xl"><SelectValue /></SelectTrigger>
                  <SelectContent>{uploadDraft.headers.map((h: string) => <SelectItem key={h} value={h}>{h}</SelectItem>)}</SelectContent>
                </Select>
              </div>
              <div className="md:col-span-3 flex gap-3">
                <Button className="rounded-2xl" onClick={finalizeUpload}>تأكيد الأعمدة ورفع البيانات</Button>
                <Button variant="outline" className="rounded-2xl" onClick={() => setUploadDraft(null)}>إلغاء</Button>
              </div>
            </CardContent>
          </Card>
        ) : null}

        {uploadReport ? (
          <Card className="rounded-2xl border">
            <CardContent className="space-y-4 p-5">
              <div className="grid gap-3 md:grid-cols-4">
                <StatCard title="اسم الملف" value={uploadReport.fileName} icon={FileSpreadsheet} />
                <StatCard title="عمود الاسم" value={uploadReport.mapping.name} icon={Briefcase} />
                <StatCard title="عمود السعر" value={uploadReport.mapping.price} icon={Briefcase} />
                <StatCard title="عمود الخصم" value={uploadReport.mapping.discount} icon={Briefcase} />
              </div>
              <div className="grid gap-3 md:grid-cols-2">
                <StatCard title="صفوف صحيحة" value={uploadReport.valid} icon={CheckCircle2} />
                <StatCard title="صفوف بها أخطاء" value={uploadReport.invalid.length} icon={XCircle} />
              </div>
            </CardContent>
          </Card>
        ) : null}
      </CardContent>
    </Card>
  );

  const CompareMapper = ({ draft, setDraft, onConfirm, title }: any) => (
    <Card className="rounded-3xl shadow-sm">
      <CardHeader>
        <CardTitle>{title}</CardTitle>
        <CardDescription>حدد الأعمدة الصحيحة قبل اعتماد الملف.</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        {!draft ? (
          <p className="text-sm text-muted-foreground">لم يتم رفع الملف بعد.</p>
        ) : (
          <>
            <div className="rounded-2xl border p-4 text-sm text-muted-foreground">{draft.fileName}</div>
            <div>
              <p className="mb-2 text-sm font-medium">عمود الاسم</p>
              <Select value={draft.mapping.name} onValueChange={(value) => setDraft((prev: any) => ({ ...prev, mapping: { ...prev.mapping, name: value } }))}>
                <SelectTrigger className="rounded-2xl"><SelectValue /></SelectTrigger>
                <SelectContent>{draft.headers.map((h: string) => <SelectItem key={h} value={h}>{h}</SelectItem>)}</SelectContent>
              </Select>
            </div>
            <div>
              <p className="mb-2 text-sm font-medium">عمود السعر</p>
              <Select value={draft.mapping.price} onValueChange={(value) => setDraft((prev: any) => ({ ...prev, mapping: { ...prev.mapping, price: value } }))}>
                <SelectTrigger className="rounded-2xl"><SelectValue /></SelectTrigger>
                <SelectContent>{draft.headers.map((h: string) => <SelectItem key={h} value={h}>{h}</SelectItem>)}</SelectContent>
              </Select>
            </div>
            <div>
              <p className="mb-2 text-sm font-medium">عمود الخصم</p>
              <Select value={draft.mapping.discount} onValueChange={(value) => setDraft((prev: any) => ({ ...prev, mapping: { ...prev.mapping, discount: value } }))}>
                <SelectTrigger className="rounded-2xl"><SelectValue /></SelectTrigger>
                <SelectContent>{draft.headers.map((h: string) => <SelectItem key={h} value={h}>{h}</SelectItem>)}</SelectContent>
              </Select>
            </div>
            <Button className="w-full rounded-2xl" onClick={onConfirm}>اعتماد الملف</Button>
          </>
        )}
      </CardContent>
    </Card>
  );

  const renderCompare = () => (
    <div className="space-y-6">
      <Alert className="rounded-2xl">
        <AlertTitle>المقارنة الذكية</AlertTitle>
        <AlertDescription>المنصة تطابق الأصناف بنسبة تشابه 80% حتى مع الاختلافات البسيطة في الكتابة.</AlertDescription>
      </Alert>

      <div className="grid gap-6 lg:grid-cols-2">
        <Card className="rounded-3xl shadow-sm">
          <CardHeader><CardTitle>رفع الملف الأول</CardTitle></CardHeader>
          <CardContent><Button className="w-full rounded-2xl" variant="outline" onClick={() => compareARef.current?.click()}><Upload className="mr-2 h-4 w-4" /> رفع الملف الأول</Button></CardContent>
        </Card>
        <Card className="rounded-3xl shadow-sm">
          <CardHeader><CardTitle>رفع الملف الثاني</CardTitle></CardHeader>
          <CardContent><Button className="w-full rounded-2xl" variant="outline" onClick={() => compareBRef.current?.click()}><Upload className="mr-2 h-4 w-4" /> رفع الملف الثاني</Button></CardContent>
        </Card>
      </div>

      <input ref={compareARef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleCompareDraft(e, "A")} />
      <input ref={compareBRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleCompareDraft(e, "B")} />

      <div className="grid gap-6 lg:grid-cols-2">
        <CompareMapper draft={compareDraftA} setDraft={setCompareDraftA} onConfirm={() => finalizeCompareFile("A")} title="إعداد الملف الأول" />
        <CompareMapper draft={compareDraftB} setDraft={setCompareDraftB} onConfirm={() => finalizeCompareFile("B")} title="إعداد الملف الثاني" />
      </div>

      <div className="flex flex-wrap justify-center gap-3">
        <Button className="rounded-2xl px-8" onClick={runComparison} disabled={!compareA || !compareB}><ArrowUpDown className="mr-2 h-4 w-4" /> ابدأ المقارنة</Button>
        <Button className="rounded-2xl px-8" variant="outline" onClick={exportComparison} disabled={!compareResult}><Download className="mr-2 h-4 w-4" /> تصدير CSV</Button>
      </div>

      {compareResult && (
        <div className="space-y-6">
          <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-6">
            <StatCard title="إجمالي ملف 1" value={compareResult.totalA} icon={FileSpreadsheet} />
            <StatCard title="إجمالي ملف 2" value={compareResult.totalB} icon={FileSpreadsheet} />
            <StatCard title="أصناف مشتركة" value={compareResult.sharedCount} icon={Users} />
            <StatCard title="الأفضل في ملف 1" value={compareResult.aBetter} icon={CheckCircle2} />
            <StatCard title="الأفضل في ملف 2" value={compareResult.bBetter} icon={CheckCircle2} />
            <StatCard title="حد التشابه" value={`${compareResult.threshold}%`} icon={Equal} />
          </div>

          <Card className="rounded-3xl shadow-sm">
            <CardHeader>
              <CardTitle className="text-2xl">نتائج المقارنة</CardTitle>
              <CardDescription>تختار لكل ملف عمود الاسم وعمود السعر وعمود الخصم، ثم تتم المقارنة الذكية بنسبة تشابه 80% حتى مع اختلاف الكتابة.</CardDescription>
            </CardHeader>
            <CardContent className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>صنف ملف 1</TableHead>
                    <TableHead>صنف ملف 2</TableHead>
                    <TableHead>التشابه</TableHead>
                    <TableHead>خصم 1</TableHead>
                    <TableHead>نهائي 1</TableHead>
                    <TableHead>خصم 2</TableHead>
                    <TableHead>نهائي 2</TableHead>
                    <TableHead>الأفضل</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {compareResult.shared.map((row: any) => (
                    <TableRow key={row.key}>
                      <TableCell className="font-medium">{row.a.productName}</TableCell>
                      <TableCell className="font-medium">{row.b.productName}</TableCell>
                      <TableCell>{row.similarity}%</TableCell>
                      <TableCell>{row.a.discount}%</TableCell>
                      <TableCell>{currency(row.a.finalPrice)} ج</TableCell>
                      <TableCell>{row.b.discount}%</TableCell>
                      <TableCell>{currency(row.b.finalPrice)} ج</TableCell>
                      <TableCell>
                        {row.winner === "Equal" ? <Badge variant="secondary" className="rounded-full">متساوي</Badge> : <Badge className="rounded-full">{row.winner}</Badge>}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </CardContent>
          </Card>
        </div>
      )}
    </div>
  );

  const renderSupplierDashboard = () => {
    const myOffers = myCompany?.offers || [];
    const topOffer = [...myOffers].sort((a: any, b: any) => b.discount - a.discount)[0];
    return (
      <div className="space-y-6">
        <SectionTitle title="لوحة المورد" desc="تابع أصنافك وارفع الملف الجديد في أي وقت." action={<Button className="rounded-2xl" onClick={() => setActiveTab("upload")}><Upload className="mr-2 h-4 w-4" />رفع ملف جديد</Button>} />
        <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
          <StatCard title="عدد الأصناف" value={myOffers.length} icon={FileSpreadsheet} />
          <StatCard title="أفضل خصم" value={topOffer ? `${topOffer.discount}%` : "0%"} icon={TrendingUp} />
          <StatCard title="الخطة الحالية" value={myCompany?.plan || "Free"} icon={Crown} />
          <StatCard title="اسم الشركة" value={myCompany?.companyName || currentUser?.companyName} icon={Store} />
        </div>
      </div>
    );
  };

  const renderAdminDashboard = () => (
    <div className="space-y-6">
      <SectionTitle title="لوحة الأدمن" desc="إدارة الموردين، العروض، ومراقبة حالة السوق." />
      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-5">
        <StatCard title="المستخدمون" value={users.length} icon={Users} />
        <StatCard title="الشركات" value={companies.length} icon={Store} />
        <StatCard title="العروض" value={allOffers.length} icon={BarChart3} />
        <StatCard title="المجموعات الذكية" value={productCount} icon={ShieldCheck} />
        <StatCard title="تنبيهات" value={uploadReport?.invalid?.length || 0} icon={Bell} hint="صفوف غير صالحة في آخر رفع" />
      </div>

      <Card className="rounded-3xl shadow-sm">
        <CardHeader>
          <CardTitle>أفضل الشركات حسب متوسط الخصم</CardTitle>
          <CardDescription>ترتيب أولي لتوضيح قيمة المنصة للمشتري.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-3">
          {marketLeaderboard.map((company: any, index: number) => (
            <div key={company.id} className="flex items-center justify-between rounded-2xl border p-4">
              <div>
                <p className="font-semibold">#{index + 1} {company.companyName}</p>
                <p className="text-sm text-muted-foreground">{company.items} صنف • متوسط خصم {company.avgDiscount}%</p>
              </div>
              <Badge className="rounded-full">{company.plan}</Badge>
            </div>
          ))}
        </CardContent>
      </Card>

      <Card className="rounded-3xl shadow-sm">
        <CardHeader>
          <CardTitle>إدارة الشركات</CardTitle>
          <CardDescription>حذف أو متابعة كل شركة مضافة داخل المنصة.</CardDescription>
        </CardHeader>
        <CardContent>
          <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-3">
            {companies.map((company: any) => (
              <Card key={company.id} className="rounded-2xl border">
                <CardContent className="flex items-center justify-between gap-3 p-4">
                  <div>
                    <p className="font-semibold">{company.companyName}</p>
                    <p className="text-sm text-muted-foreground">{company.offers.length} صنف • {company.plan}</p>
                  </div>
                  <Button variant="ghost" size="icon" onClick={() => removeCompany(company.id)}>
                    <Trash2 className="h-4 w-4" />
                  </Button>
                </CardContent>
              </Card>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="mx-auto max-w-7xl space-y-8">
        {!currentUser ? (
          renderAuth()
        ) : (
          <>
            {renderTopBar()}
            <Tabs value={activeTab} onValueChange={setActiveTab} className="space-y-6">
              <TabsList className="grid w-full grid-cols-4 rounded-2xl">
                <TabsTrigger value="market" className="rounded-2xl">بحث السوق</TabsTrigger>
                <TabsTrigger value="upload" className="rounded-2xl">رفع ملف</TabsTrigger>
                <TabsTrigger value="compare" className="rounded-2xl">قارن ملفين</TabsTrigger>
                <TabsTrigger value={currentUser.role === "admin" ? "admin" : "dashboard"} className="rounded-2xl">
                  {currentUser.role === "admin" ? <><LayoutDashboard className="mr-2 h-4 w-4" />الأدمن</> : <><LayoutDashboard className="mr-2 h-4 w-4" />لوحتي</>}
                </TabsTrigger>
              </TabsList>
              <TabsContent value="market">{renderMarket()}</TabsContent>
              <TabsContent value="upload">{renderUpload()}</TabsContent>
              <TabsContent value="compare">{renderCompare()}</TabsContent>
              <TabsContent value="dashboard">{renderSupplierDashboard()}</TabsContent>
              <TabsContent value="admin">{renderAdminDashboard()}</TabsContent>
            </Tabs>
          </>
        )}
      </div>
    </div>
  );
}
