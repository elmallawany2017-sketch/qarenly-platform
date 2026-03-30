"use client";

import React, { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Offer = {
  productName: string;
  normalizedName: string;
  price: number;
  discount: number;
  finalPrice: number;
};

type Company = {
  id: string;
  companyName: string;
  email: string;
  plan: string;
  code?: string;
  uploadedAt?: string;
  offers: Offer[];
};

type User = {
  id: string;
  name: string;
  email: string;
  password: string;
  role: "admin" | "supplier";
  companyName: string;
};

type RawRow = Record<string, unknown>;

type UploadDraft = {
  rows: RawRow[];
  headers: string[];
  fileName: string;
  mapping: {
    name: string;
    price: string;
    discount: string;
  };
};

const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;
const MATCH_THRESHOLD = 0.8;

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

const currency = (n: number | string) => {
  const num = Number(n || 0);
  return new Intl.NumberFormat("en-EG", {
    maximumFractionDigits: 2,
  }).format(num);
};

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
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }

  return dp[m][n];
};

const similarityScore = (a: string, b: string) => {
  const x = normalizeName(a);
  const y = normalizeName(b);

  if (!x || !y) return 0;
  if (x === y) return 1;

  if (x.includes(y) || y.includes(x)) {
    return Math.min(x.length, y.length) / Math.max(x.length, y.length);
  }

  const distance = levenshtein(x, y);
  const maxLen = Math.max(x.length, y.length);
  return maxLen === 0 ? 1 : 1 - distance / maxLen;
};

const bestMatch = (target: string, candidates: string[], threshold = MATCH_THRESHOLD) => {
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
    name: ["name", "product name", "اسم الصنف", "الصنف", "اسم المنتج", "item", "description"],
    price: ["price", "السعر", "سعر", "unit price", "buy price", "cost"],
    discount: ["discount", "الخصم", "discount %", "خصم", "نسبه الخصم", "disc"],
  }[kind];

  for (const pattern of patterns) {
    const found = normalized.find(
      (h) =>
        h.key === normalizeName(pattern) ||
        h.key.includes(normalizeName(pattern))
    );
    if (found) return found.raw;
  }

  return headers[0] || "";
};

const readExcelRaw = async (file: File) => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" }) as RawRow[];
  const headers = rows.length ? Object.keys(rows[0]) : [];
  return { rows, headers, fileName: file.name };
};

const mapRowsFromSelection = (
  rows: RawRow[],
  mapping: { name: string; price: string; discount: string }
) => {
  return rows.map((row, index) => {
    const productName = String(row[mapping.name] ?? "").trim();
    const price = Number(row[mapping.price]);
    const discount = Number(row[mapping.discount]);
    const errors: string[] = [];

    if (!productName) errors.push("Missing product name");
    if (Number.isNaN(price) || price <= 0) errors.push("Invalid price");
    if (Number.isNaN(discount) || discount < 0 || discount > 100) {
      errors.push("Invalid discount");
    }

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

const seedCompaniesBase = [
  {
    companyName: "Alpha Pharma",
    email: "alpha@qarenly.com",
    plan: "Pro",
    offers: [
      { productName: "Panadol Extra", price: 100, discount: 20 },
      { productName: "Augmentin 1g", price: 180, discount: 12 },
      { productName: "Cetal 500", price: 30, discount: 5 },
    ],
  },
  {
    companyName: "Trust Med",
    email: "trust@qarenly.com",
    plan: "Business",
    offers: [
      { productName: "Panadol Exra", price: 98, discount: 15 },
      { productName: "Augmentin 1 gm", price: 175, discount: 8 },
      { productName: "Brufen 400", price: 48, discount: 10 },
    ],
  },
  {
    companyName: "Market Plus",
    email: "market@qarenly.com",
    plan: "Free",
    offers: [
      { productName: "Panadol Extra", price: 101, discount: 22 },
      { productName: "Cetal500", price: 31, discount: 8 },
      { productName: "Brufen 400", price: 47, discount: 5 },
    ],
  },
];

const seedCompanies: Company[] = seedCompaniesBase.map((c, i) => ({
  id: crypto.randomUUID(),
  companyName: c.companyName,
  email: c.email,
  plan: c.plan,
  code: `C${String(i + 1).padStart(3, "0")}`,
  uploadedAt: new Date().toISOString(),
  offers: c.offers.map((o) => ({
    productName: o.productName,
    normalizedName: normalizeName(o.productName),
    price: o.price,
    discount: o.discount,
    finalPrice: calcFinalPrice(o.price, o.discount),
  })),
}));

const seedUsers: User[] = [
  {
    id: "admin-1",
    name: "Admin",
    email: "admin@qarenly.com",
    password: "123456",
    role: "admin",
    companyName: "Qarenly",
  },
  {
    id: "supplier-1",
    name: "Alpha Pharma",
    email: "alpha@qarenly.com",
    password: "123456",
    role: "supplier",
    companyName: "Alpha Pharma",
  },
  {
    id: "supplier-2",
    name: "Trust Med",
    email: "trust@qarenly.com",
    password: "123456",
    role: "supplier",
    companyName: "Trust Med",
  },
];

function Card({
  children,
  className = "",
}: {
  children: React.ReactNode;
  className?: string;
}) {
  return (
    <div className={`rounded-2xl border border-slate-200 bg-white shadow-sm ${className}`}>
      {children}
    </div>
  );
}

function StatBox({
  title,
  value,
  hint,
}: {
  title: string;
  value: string | number;
  hint?: string;
}) {
  return (
    <Card>
      <div className="p-5">
        <div className="text-sm text-slate-500">{title}</div>
        <div className="mt-2 text-2xl font-bold text-slate-900">{value}</div>
        {hint ? <div className="mt-1 text-xs text-slate-400">{hint}</div> : null}
      </div>
    </Card>
  );
}

function SectionTitle({
  title,
  desc,
  action,
}: {
  title: string;
  desc?: string;
  action?: React.ReactNode;
}) {
  return (
    <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
      <div>
        <h2 className="text-2xl font-bold text-slate-900">{title}</h2>
        {desc ? <p className="mt-1 text-sm text-slate-500">{desc}</p> : null}
      </div>
      {action}
    </div>
  );
}

export default function Page() {
  const [companies, setCompanies] = useState<Company[]>(seedCompanies);
  const [users, setUsers] = useState<User[]>(seedUsers);

  const [authMode, setAuthMode] = useState<"login" | "register">("login");
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [activeTab, setActiveTab] = useState("market");

  const [loginForm, setLoginForm] = useState({
    email: "admin@qarenly.com",
    password: "123456",
  });

  const [registerForm, setRegisterForm] = useState({
    name: "",
    email: "",
    password: "",
    companyName: "",
  });

  const [authError, setAuthError] = useState("");
  const [search, setSearch] = useState("");
  const [sortBy, setSortBy] = useState("discount");
  const [filterCompany, setFilterCompany] = useState("all");

  const [companyName, setCompanyName] = useState("");
  const [companyCode, setCompanyCode] = useState("");
  const [uploadReport, setUploadReport] = useState<any>(null);
  const [productSuggestions, setProductSuggestions] = useState<string[]>([]);

  const [uploadDraft, setUploadDraft] = useState<UploadDraft | null>(null);
  const [compareDraftA, setCompareDraftA] = useState<UploadDraft | null>(null);
  const [compareDraftB, setCompareDraftB] = useState<UploadDraft | null>(null);
  const [compareA, setCompareA] = useState<any>(null);
  const [compareB, setCompareB] = useState<any>(null);
  const [compareResult, setCompareResult] = useState<any>(null);

  const companyFileRef = useRef<HTMLInputElement>(null);
  const compareARef = useRef<HTMLInputElement>(null);
  const compareBRef = useRef<HTMLInputElement>(null);

  const validCompanies = useMemo(() => {
    const now = Date.now();
    return companies.filter((company) => {
      if (!company.uploadedAt) return true;
      const age = now - new Date(company.uploadedAt).getTime();
      return age <= SEVEN_DAYS_MS;
    });
  }, [companies]);

  const allOffers = useMemo(() => {
    return validCompanies.flatMap((company) =>
      company.offers.map((offer) => ({
        ...offer,
        companyName: company.companyName,
        companyId: company.id,
        plan: company.plan,
        code: company.code,
      }))
    );
  }, [validCompanies]);

  const groupedResults = useMemo(() => {
    const query = normalizeName(search);

    let data = allOffers;

    if (query) {
      data = data.filter(
        (offer) =>
          offer.normalizedName.includes(query) ||
          normalizeName(offer.companyName).includes(query)
      );
    }

    if (filterCompany !== "all") {
      data = data.filter((offer) => offer.companyId === filterCompany);
    }

    const groups: Array<{
      key: string;
      anchor: string;
      displayName: string;
      rows: any[];
    }> = [];

    for (const row of data) {
      let group = groups.find(
        (g) => similarityScore(g.anchor, row.productName) >= MATCH_THRESHOLD
      );

      if (!group) {
        group = {
          key: crypto.randomUUID(),
          anchor: row.productName,
          displayName: row.productName,
          rows: [],
        };
        groups.push(group);
      }

      group.rows.push(row);
    }

    groups.forEach((group) => {
      group.rows.sort((a, b) => {
        if (sortBy === "discount") {
          return b.discount - a.discount || a.finalPrice - b.finalPrice;
        }
        if (sortBy === "final") {
          return a.finalPrice - b.finalPrice || b.discount - a.discount;
        }
        return a.productName.localeCompare(b.productName);
      });
    });

    groups.sort((a, b) => a.displayName.localeCompare(b.displayName));
    return groups;
  }, [allOffers, filterCompany, search, sortBy]);

  const productCount = groupedResults.length;

  const avgDiscount = allOffers.length
    ? (
        allOffers.reduce((sum, row) => sum + row.discount, 0) / allOffers.length
      ).toFixed(1)
    : "0";

  const myCompany = useMemo(() => {
    if (!currentUser || currentUser.role !== "supplier") return null;
    return validCompanies.find(
      (c) => normalizeName(c.companyName) === normalizeName(currentUser.companyName)
    );
  }, [currentUser, validCompanies]);

  const marketLeaderboard = useMemo(() => {
    return validCompanies
      .map((company) => {
        const offers = company.offers || [];
        const avg = offers.length
          ? offers.reduce((sum, row) => sum + row.discount, 0) / offers.length
          : 0;
        return {
          ...company,
          avgDiscount: avg.toFixed(1),
          items: offers.length,
        };
      })
      .sort((a, b) => Number(b.avgDiscount) - Number(a.avgDiscount));
  }, [validCompanies]);

  const login = () => {
    const found = users.find(
      (u) => u.email === loginForm.email && u.password === loginForm.password
    );

    if (!found) {
      setAuthError("بيانات الدخول غير صحيحة");
      return;
    }

    setAuthError("");
    setCurrentUser(found);
    setActiveTab(found.role === "admin" ? "admin" : "dashboard");
  };

  const register = () => {
    if (
      !registerForm.name ||
      !registerForm.email ||
      !registerForm.password ||
      !registerForm.companyName
    ) {
      setAuthError("اكمل كل البيانات أولًا");
      return;
    }

    if (users.some((u) => u.email === registerForm.email)) {
      setAuthError("هذا البريد مستخدم بالفعل");
      return;
    }

    const newUser: User = {
      id: crypto.randomUUID(),
      name: registerForm.name,
      email: registerForm.email,
      password: registerForm.password,
      role: "supplier",
      companyName: registerForm.companyName,
    };

    setUsers((prev) => [...prev, newUser]);
    setCompanies((prev) => [
      ...prev,
      {
        id: crypto.randomUUID(),
        companyName: registerForm.companyName,
        email: registerForm.email,
        plan: "Free",
        code: `C${String(prev.length + 1).padStart(3, "0")}`,
        uploadedAt: new Date().toISOString(),
        offers: [],
      },
    ]);
    setCurrentUser(newUser);
    setAuthError("");
    setActiveTab("dashboard");
    setRegisterForm({ name: "", email: "", password: "", companyName: "" });
  };

  const logout = () => {
    setCurrentUser(null);
    setActiveTab("market");
  };

  const handleCompanyDraft = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
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

    const resolvedCompanyName =
      currentUser?.role === "supplier"
        ? currentUser.companyName
        : companyName.trim() || uploadDraft.fileName.replace(/\.[^.]+$/, "");

    const mappedOffers = validRows.map(({ rowNumber, errors, valid, ...rest }) => rest);

    setCompanies((prev) => {
      const existingIndex = prev.findIndex(
        (c) =>
          normalizeName(c.companyName) === normalizeName(resolvedCompanyName) ||
          (!!companyCode && c.code === companyCode)
      );

      if (existingIndex >= 0) {
        const clone = [...prev];
        clone[existingIndex] = {
          ...clone[existingIndex],
          companyName: resolvedCompanyName,
          code: clone[existingIndex].code || companyCode,
          offers: mappedOffers,
          uploadedAt: new Date().toISOString(),
        };
        return clone;
      }

      return [
        ...prev,
        {
          id: crypto.randomUUID(),
          companyName: resolvedCompanyName,
          email: currentUser?.email || "",
          plan: "Free",
          code: companyCode || `C${String(prev.length + 1).padStart(3, "0")}`,
          uploadedAt: new Date().toISOString(),
          offers: mappedOffers,
        },
      ];
    });

    setUploadReport({
      fileName: uploadDraft.fileName,
      companyName: resolvedCompanyName,
      total: rows.length,
      valid: validRows.length,
      invalid: invalidRows,
      mapping: uploadDraft.mapping,
      expiresIn: "7 days",
      companyCode:
        companyCode ||
        validCompanies.find(
          (c) => normalizeName(c.companyName) === normalizeName(resolvedCompanyName)
        )?.code ||
        "سيتم التكويد",
    });

    setUploadDraft(null);
    setCompanyName("");
    setCompanyCode("");
  };

  const handleCompareDraft = async (
    event: React.ChangeEvent<HTMLInputElement>,
    which: "A" | "B"
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const draft = await readExcelRaw(file);
    const mapping = {
      name: guessColumn(draft.headers, "name"),
      price: guessColumn(draft.headers, "price"),
      discount: guessColumn(draft.headers, "discount"),
    };

    const payload = {
      ...draft,
      mapping,
      fileName: draft.fileName,
    };

    if (which === "A") setCompareDraftA(payload);
    else setCompareDraftB(payload);

    event.target.value = "";
  };

  const finalizeCompareFile = (which: "A" | "B") => {
    const draft = which === "A" ? compareDraftA : compareDraftB;
    if (!draft) return;

    const rows = mapRowsFromSelection(draft.rows, draft.mapping)
      .filter((r) => r.valid)
      .map(({ rowNumber, errors, valid, ...rest }) => rest);

    const payload = {
      name: which === "A" ? "ملف 1" : "ملف 2",
      rawName: draft.fileName,
      rows,
      mapping: draft.mapping,
    };

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
      const match = bestMatch(
        a.productName,
        bKeys.filter((name: string) => !usedB.has(name)),
        MATCH_THRESHOLD
      );

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
        a,
        b,
        similarity: +(match.score * 100).toFixed(0),
        winner,
      });
    }

    const onlyB = compareB.rows.filter((row: any) => !usedB.has(row.productName));

    setCompareResult({
      shared,
      onlyA,
      onlyB,
      aBetter: shared.filter((r) => r.winner === compareA.name).length,
      bBetter: shared.filter((r) => r.winner === compareB.name).length,
      equal: shared.filter((r) => r.winner === "Equal").length,
      totalA: compareA.rows.length,
      totalB: compareB.rows.length,
      sharedCount: shared.length,
      threshold: 80,
    });
  };

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

  const removeCompany = (id: string) => {
    setCompanies((prev) => prev.filter((c) => c.id !== id));
  };

  const onSearchChange = (value: string) => {
    setSearch(value);

    if (value.trim().length >= 3) {
      const suggestions = Array.from(
        new Set(
          allOffers
            .filter((offer) =>
              normalizeName(offer.productName).includes(normalizeName(value))
            )
            .map((offer) => offer.productName)
        )
      ).slice(0, 8);

      setProductSuggestions(suggestions);
    } else {
      setProductSuggestions([]);
    }
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-gradient-to-b from-slate-50 to-white p-4 md:p-8">
        <div className="mx-auto grid max-w-7xl gap-6 lg:grid-cols-[1.15fr_.85fr]">
          <Card className="border-0 bg-gradient-to-br from-slate-900 via-slate-800 to-slate-700 text-white shadow-xl">
            <div className="p-8 md:p-10">
              <div className="mb-4 inline-block rounded-full bg-white/10 px-4 py-1 text-sm">
                Qarenly Pro
              </div>
              <h1 className="text-3xl font-bold tracking-tight md:text-5xl">
                منصة احترافية لمقارنة الخصومات والأسعار
              </h1>
              <p className="mt-4 max-w-2xl text-base text-slate-200 md:text-lg">
                تكويد العملاء بواسطة الأدمن، رفع مرن للإكسيل باختيار الأعمدة،
                استبدال تلقائي لملف العميل القديم، حذف تلقائي بعد 7 أيام،
                واقتراحات أصناف بمجرد كتابة أول 3 حروف.
              </p>

              <div className="mt-8 grid gap-4 sm:grid-cols-3">
                <StatBox title="الشركات" value={validCompanies.length} hint="فعالة خلال 7 أيام" />
                <StatBox title="العروض" value={allOffers.length} hint="إجمالي العروض الفعالة" />
                <StatBox title="متوسط الخصم" value={`${avgDiscount}%`} hint="على مستوى السوق" />
              </div>
            </div>
          </Card>

          <Card className="shadow-xl">
            <div className="p-6">
              <h2 className="text-2xl font-bold text-slate-900">
                {authMode === "login" ? "تسجيل الدخول" : "إنشاء حساب مورد"}
              </h2>
              <p className="mt-2 text-sm text-slate-500">
                حساب الأدمن يدير الأكواد والعملاء، والمورد يرفع ملفه فقط.
              </p>

              <div className="mt-5 space-y-4">
                {authMode === "login" ? (
                  <>
                    <input
                      className="w-full rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                      placeholder="البريد الإلكتروني"
                      value={loginForm.email}
                      onChange={(e) =>
                        setLoginForm((p) => ({ ...p, email: e.target.value }))
                      }
                    />
                    <input
                      className="w-full rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                      type="password"
                      placeholder="كلمة المرور"
                      value={loginForm.password}
                      onChange={(e) =>
                        setLoginForm((p) => ({ ...p, password: e.target.value }))
                      }
                    />
                    <button
                      className="w-full rounded-xl bg-slate-900 p-3 font-medium text-white hover:bg-slate-800"
                      onClick={login}
                    >
                      دخول
                    </button>
                    <div className="rounded-xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-600">
                      <p>حساب الأدمن: admin@qarenly.com</p>
                      <p>كلمة المرور: 123456</p>
                    </div>
                  </>
                ) : (
                  <>
                    <input
                      className="w-full rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                      placeholder="الاسم"
                      value={registerForm.name}
                      onChange={(e) =>
                        setRegisterForm((p) => ({ ...p, name: e.target.value }))
                      }
                    />
                    <input
                      className="w-full rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                      placeholder="اسم الشركة"
                      value={registerForm.companyName}
                      onChange={(e) =>
                        setRegisterForm((p) => ({
                          ...p,
                          companyName: e.target.value,
                        }))
                      }
                    />
                    <input
                      className="w-full rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                      placeholder="البريد الإلكتروني"
                      value={registerForm.email}
                      onChange={(e) =>
                        setRegisterForm((p) => ({ ...p, email: e.target.value }))
                      }
                    />
                    <input
                      className="w-full rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                      type="password"
                      placeholder="كلمة المرور"
                      value={registerForm.password}
                      onChange={(e) =>
                        setRegisterForm((p) => ({ ...p, password: e.target.value }))
                      }
                    />
                    <button
                      className="w-full rounded-xl bg-slate-900 p-3 font-medium text-white hover:bg-slate-800"
                      onClick={register}
                    >
                      إنشاء الحساب
                    </button>
                  </>
                )}

                {authError ? (
                  <div className="rounded-xl border border-red-200 bg-red-50 p-3 text-sm text-red-700">
                    {authError}
                  </div>
                ) : null}

                <button
                  className="w-full rounded-xl border border-slate-300 p-3 hover:bg-slate-50"
                  onClick={() => {
                    setAuthMode(authMode === "login" ? "register" : "login");
                    setAuthError("");
                  }}
                >
                  {authMode === "login"
                    ? "إنشاء حساب جديد"
                    : "لديك حساب بالفعل؟ سجل دخول"}
                </button>
              </div>
            </div>
          </Card>
        </div>
      </div>
    );
  }

  const myOffers = myCompany?.offers || [];
  const topOffer = [...myOffers].sort((a, b) => b.discount - a.discount)[0];

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 to-white p-4 md:p-8">
      <div className="mx-auto max-w-7xl space-y-8">
        <Card className="border-0 shadow-lg">
          <div className="flex flex-col gap-5 p-6 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <div className="rounded-full bg-slate-900 px-4 py-1 text-sm text-white">
                  {currentUser.role === "admin" ? "Admin" : "Supplier"}
                </div>
                <div className="rounded-full bg-slate-100 px-4 py-1 text-sm text-slate-700">
                  {currentUser.companyName}
                </div>
              </div>
              <h1 className="mt-3 text-3xl font-bold text-slate-900">
                مرحبًا، {currentUser.name}
              </h1>
              <p className="mt-1 text-sm text-slate-500">
                واجهة أكثر احترافية، مع تكويد العملاء بواسطة الأدمن، واستبدال تلقائي
                لملف نفس العميل، ومسح تلقائي بعد 7 أيام، واقتراحات أصناف عند كتابة
                أول 3 حروف.
              </p>
            </div>
            <button
              className="rounded-xl border border-slate-300 px-4 py-2 hover:bg-slate-50"
              onClick={logout}
            >
              تسجيل خروج
            </button>
          </div>
        </Card>

        <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
          <button
            onClick={() => setActiveTab("market")}
            className={`rounded-xl p-3 font-medium ${
              activeTab === "market"
                ? "bg-slate-900 text-white"
                : "border border-slate-300 bg-white text-slate-700"
            }`}
          >
            بحث السوق
          </button>
          <button
            onClick={() => setActiveTab("upload")}
            className={`rounded-xl p-3 font-medium ${
              activeTab === "upload"
                ? "bg-slate-900 text-white"
                : "border border-slate-300 bg-white text-slate-700"
            }`}
          >
            رفع ملف
          </button>
          <button
            onClick={() => setActiveTab("compare")}
            className={`rounded-xl p-3 font-medium ${
              activeTab === "compare"
                ? "bg-slate-900 text-white"
                : "border border-slate-300 bg-white text-slate-700"
            }`}
          >
            قارن ملفين
          </button>
          <button
            onClick={() =>
              setActiveTab(currentUser.role === "admin" ? "admin" : "dashboard")
            }
            className={`rounded-xl p-3 font-medium ${
              activeTab === (currentUser.role === "admin" ? "admin" : "dashboard")
                ? "bg-slate-900 text-white"
                : "border border-slate-300 bg-white text-slate-700"
            }`}
          >
            {currentUser.role === "admin" ? "الأدمن" : "لوحتي"}
          </button>
        </div>

        {activeTab === "market" && (
          <div className="space-y-6">
            <Card>
              <div className="p-6">
                <SectionTitle
                  title="محرك مقارنة السوق"
                  desc="عند البحث عن أي صنف ستظهر كل الشركات أو العملاء العارضين له، مع اسم العميل والكود والخصم والسعر والسعر النهائي، مرتبين حسب أعلى خصم."
                />

                <div className="mt-4 grid gap-4 lg:grid-cols-[1fr_220px_260px]">
                  <div className="relative">
                    <input
                      className="w-full rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                      value={search}
                      onChange={(e) => onSearchChange(e.target.value)}
                      placeholder="ابحث باسم الصنف أو الشركة..."
                    />
                    {productSuggestions.length > 0 && (
                      <div className="absolute z-10 mt-1 w-full rounded-xl border border-slate-200 bg-white shadow-lg">
                        {productSuggestions.map((suggestion) => (
                          <button
                            key={suggestion}
                            type="button"
                            className="block w-full border-b border-slate-100 px-4 py-2 text-right text-sm hover:bg-slate-50"
                            onClick={() => {
                              setSearch(suggestion);
                              setProductSuggestions([]);
                            }}
                          >
                            {suggestion}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>

                  <select
                    className="rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                    value={sortBy}
                    onChange={(e) => setSortBy(e.target.value)}
                  >
                    <option value="discount">أعلى خصم</option>
                    <option value="final">أقل سعر نهائي</option>
                    <option value="name">اسم الصنف</option>
                  </select>

                  <select
                    className="rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                    value={filterCompany}
                    onChange={(e) => setFilterCompany(e.target.value)}
                  >
                    <option value="all">كل الشركات</option>
                    {validCompanies.map((company) => (
                      <option key={company.id} value={company.id}>
                        {company.companyName}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </Card>

            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              <StatBox title="إجمالي الشركات" value={validCompanies.length} hint="فعالة خلال 7 أيام" />
              <StatBox title="إجمالي العروض" value={allOffers.length} />
              <StatBox title="مجموعات الأصناف" value={productCount} hint="بالتجميع الذكي 80%" />
              <StatBox title="متوسط الخصم" value={`${avgDiscount}%`} />
            </div>

            <div className="space-y-5">
              {groupedResults.length === 0 ? (
                <Card>
                  <div className="p-8 text-center text-slate-500">لا توجد نتائج مطابقة.</div>
                </Card>
              ) : (
                groupedResults.map((group) => (
                  <Card key={group.key}>
                    <div className="p-6">
                      <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
                        <div>
                          <h3 className="text-xl font-bold text-slate-900">{group.displayName}</h3>
                          <p className="text-sm text-slate-500">
                            {group.rows.length} عروض متقاربة
                          </p>
                        </div>
                        <div className="rounded-full bg-slate-100 px-4 py-1 text-sm text-slate-700">
                          أفضل عرض:{" "}
                          {sortBy === "final"
                            ? `${currency(group.rows[0].finalPrice)} ج`
                            : `${group.rows[0].discount}%`}
                        </div>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="min-w-full border-collapse">
                          <thead>
                            <tr className="border-b border-slate-200 text-right text-sm text-slate-500">
                              <th className="p-3">الترتيب</th>
                              <th className="p-3">الشركة / العميل</th>
                              <th className="p-3">الكود</th>
                              <th className="p-3">اسم الصنف</th>
                              <th className="p-3">الخصم</th>
                              <th className="p-3">السعر</th>
                              <th className="p-3">السعر النهائي</th>
                            </tr>
                          </thead>
                          <tbody>
                            {group.rows.map((row: any, index: number) => (
                              <tr key={`${row.companyId}-${index}`} className="border-b border-slate-100">
                                <td className="p-3">{index + 1}</td>
                                <td className="p-3 font-medium text-slate-900">{row.companyName}</td>
                                <td className="p-3">{row.code || "—"}</td>
                                <td className="p-3">{row.productName}</td>
                                <td className="p-3">
                                  <span
                                    className={`rounded-full px-3 py-1 text-sm ${
                                      index === 0
                                        ? "bg-emerald-100 text-emerald-700"
                                        : "bg-slate-100 text-slate-700"
                                    }`}
                                  >
                                    {row.discount}%
                                  </span>
                                </td>
                                <td className="p-3">{currency(row.price)} ج</td>
                                <td className="p-3">{currency(row.finalPrice)} ج</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </Card>
                ))
              )}
            </div>
          </div>
        )}

        {activeTab === "upload" && (
          <Card>
            <div className="p-6 space-y-5">
              <SectionTitle
                title="رفع ملف شركة"
                desc="بعد رفع الملف ستختار بنفسك: عمود الاسم + عمود السعر + عمود الخصم. تكويد العميل يتم بواسطة الأدمن، ورفع ملف جديد لنفس العميل يستبدل القديم تلقائيًا، وكل ملف صلاحيته 7 أيام فقط."
              />

              {currentUser.role !== "supplier" ? (
                <div className="grid gap-4 md:grid-cols-[1fr_220px_auto]">
                  <input
                    className="rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                    value={companyName}
                    onChange={(e) => setCompanyName(e.target.value)}
                    placeholder="اسم الشركة أو العميل"
                  />
                  <input
                    className="rounded-xl border border-slate-300 p-3 outline-none focus:border-slate-700"
                    value={companyCode}
                    onChange={(e) => setCompanyCode(e.target.value)}
                    placeholder="كود العميل (يحدده الأدمن)"
                  />
                  <button
                    className="rounded-xl bg-slate-900 px-4 py-3 text-white hover:bg-slate-800"
                    onClick={() => companyFileRef.current?.click()}
                  >
                    اختر ملف Excel
                  </button>
                </div>
              ) : (
                <div className="rounded-xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-600">
                  سيتم ربط الملف تلقائيًا بشركتك: <strong>{currentUser.companyName}</strong>
                </div>
              )}

              {currentUser.role === "supplier" && (
                <button
                  className="rounded-xl bg-slate-900 px-4 py-3 text-white hover:bg-slate-800"
                  onClick={() => companyFileRef.current?.click()}
                >
                  رفع ملف Excel
                </button>
              )}

              <input
                ref={companyFileRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={handleCompanyDraft}
              />

              {uploadDraft && (
                <Card className="bg-slate-50">
                  <div className="grid gap-4 p-5 md:grid-cols-3">
                    <div>
                      <p className="mb-2 text-sm font-medium text-slate-700">اختار عمود الاسم</p>
                      <select
                        className="w-full rounded-xl border border-slate-300 p-3"
                        value={uploadDraft.mapping.name}
                        onChange={(e) =>
                          setUploadDraft((prev) =>
                            prev
                              ? {
                                  ...prev,
                                  mapping: { ...prev.mapping, name: e.target.value },
                                }
                              : prev
                          )
                        }
                      >
                        {uploadDraft.headers.map((h) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <p className="mb-2 text-sm font-medium text-slate-700">اختار عمود السعر</p>
                      <select
                        className="w-full rounded-xl border border-slate-300 p-3"
                        value={uploadDraft.mapping.price}
                        onChange={(e) =>
                          setUploadDraft((prev) =>
                            prev
                              ? {
                                  ...prev,
                                  mapping: { ...prev.mapping, price: e.target.value },
                                }
                              : prev
                          )
                        }
                      >
                        {uploadDraft.headers.map((h) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <p className="mb-2 text-sm font-medium text-slate-700">اختار عمود الخصم</p>
                      <select
                        className="w-full rounded-xl border border-slate-300 p-3"
                        value={uploadDraft.mapping.discount}
                        onChange={(e) =>
                          setUploadDraft((prev) =>
                            prev
                              ? {
                                  ...prev,
                                  mapping: { ...prev.mapping, discount: e.target.value },
                                }
                              : prev
                          )
                        }
                      >
                        {uploadDraft.headers.map((h) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div className="flex gap-3 md:col-span-3">
                      <button
                        className="rounded-xl bg-slate-900 px-4 py-3 text-white hover:bg-slate-800"
                        onClick={finalizeUpload}
                      >
                        تأكيد الأعمدة ورفع البيانات
                      </button>
                      <button
                        className="rounded-xl border border-slate-300 px-4 py-3 hover:bg-white"
                        onClick={() => setUploadDraft(null)}
                      >
                        إلغاء
                      </button>
                    </div>
                  </div>
                </Card>
              )}

              {uploadReport && (
                <Card>
                  <div className="space-y-4 p-5">
                    <div className="grid gap-3 md:grid-cols-6">
                      <StatBox title="اسم الملف" value={uploadReport.fileName} />
                      <StatBox title="العميل" value={uploadReport.companyName} />
                      <StatBox title="الكود" value={uploadReport.companyCode} />
                      <StatBox title="عمود الاسم" value={uploadReport.mapping.name} />
                      <StatBox title="عمود السعر" value={uploadReport.mapping.price} />
                      <StatBox title="عمود الخصم" value={uploadReport.mapping.discount} />
                    </div>
                    <div className="grid gap-3 md:grid-cols-3">
                      <StatBox title="صفوف صحيحة" value={uploadReport.valid} />
                      <StatBox title="صفوف بها أخطاء" value={uploadReport.invalid.length} />
                      <StatBox title="الصلاحية" value={uploadReport.expiresIn} />
                    </div>
                  </div>
                </Card>
              )}
            </div>
          </Card>
        )}

        {activeTab === "compare" && (
          <div className="space-y-6">
            <Card className="bg-slate-50">
              <div className="p-5 text-sm text-slate-600">
                المنصة تطابق الأصناف بنسبة تشابه 80% حتى مع الاختلافات البسيطة في الكتابة.
                وفي كل ملف تختار عمود الاسم وعمود السعر وعمود الخصم قبل اعتماد المقارنة.
              </div>
            </Card>

            <div className="grid gap-6 lg:grid-cols-2">
              <Card>
                <div className="p-6">
                  <h3 className="mb-4 text-xl font-bold text-slate-900">رفع الملف الأول</h3>
                  <button
                    className="w-full rounded-xl border border-slate-300 p-3 hover:bg-slate-50"
                    onClick={() => compareARef.current?.click()}
                  >
                    رفع الملف الأول
                  </button>
                </div>
              </Card>

              <Card>
                <div className="p-6">
                  <h3 className="mb-4 text-xl font-bold text-slate-900">رفع الملف الثاني</h3>
                  <button
                    className="w-full rounded-xl border border-slate-300 p-3 hover:bg-slate-50"
                    onClick={() => compareBRef.current?.click()}
                  >
                    رفع الملف الثاني
                  </button>
                </div>
              </Card>
            </div>

            <input
              ref={compareARef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => handleCompareDraft(e, "A")}
            />
            <input
              ref={compareBRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(e) => handleCompareDraft(e, "B")}
            />

            <div className="grid gap-6 lg:grid-cols-2">
              <Card>
                <div className="space-y-4 p-6">
                  <h3 className="text-xl font-bold text-slate-900">إعداد الملف الأول</h3>
                  <p className="text-sm text-slate-500">
                    حدد عمود الاسم وعمود السعر وعمود الخصم قبل اعتماد الملف.
                  </p>

                  {!compareDraftA ? (
                    <p className="text-sm text-slate-500">لم يتم رفع الملف بعد.</p>
                  ) : (
                    <>
                      <div className="rounded-xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-500">
                        {compareDraftA.fileName}
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium text-slate-700">عمود الاسم</p>
                        <select
                          className="w-full rounded-xl border border-slate-300 p-3"
                          value={compareDraftA.mapping.name}
                          onChange={(e) =>
                            setCompareDraftA((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    mapping: { ...prev.mapping, name: e.target.value },
                                  }
                                : prev
                            )
                          }
                        >
                          {compareDraftA.headers.map((h: string) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium text-slate-700">عمود السعر</p>
                        <select
                          className="w-full rounded-xl border border-slate-300 p-3"
                          value={compareDraftA.mapping.price}
                          onChange={(e) =>
                            setCompareDraftA((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    mapping: { ...prev.mapping, price: e.target.value },
                                  }
                                : prev
                            )
                          }
                        >
                          {compareDraftA.headers.map((h: string) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium text-slate-700">عمود الخصم</p>
                        <select
                          className="w-full rounded-xl border border-slate-300 p-3"
                          value={compareDraftA.mapping.discount}
                          onChange={(e) =>
                            setCompareDraftA((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    mapping: { ...prev.mapping, discount: e.target.value },
                                  }
                                : prev
                            )
                          }
                        >
                          {compareDraftA.headers.map((h: string) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </div>

                      <button
                        className="w-full rounded-xl bg-slate-900 p-3 text-white hover:bg-slate-800"
                        onClick={() => finalizeCompareFile("A")}
                      >
                        اعتماد الملف
                      </button>
                    </>
                  )}
                </div>
              </Card>

              <Card>
                <div className="space-y-4 p-6">
                  <h3 className="text-xl font-bold text-slate-900">إعداد الملف الثاني</h3>
                  <p className="text-sm text-slate-500">
                    حدد عمود الاسم وعمود السعر وعمود الخصم قبل اعتماد الملف.
                  </p>

                  {!compareDraftB ? (
                    <p className="text-sm text-slate-500">لم يتم رفع الملف بعد.</p>
                  ) : (
                    <>
                      <div className="rounded-xl border border-slate-200 bg-slate-50 p-4 text-sm text-slate-500">
                        {compareDraftB.fileName}
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium text-slate-700">عمود الاسم</p>
                        <select
                          className="w-full rounded-xl border border-slate-300 p-3"
                          value={compareDraftB.mapping.name}
                          onChange={(e) =>
                            setCompareDraftB((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    mapping: { ...prev.mapping, name: e.target.value },
                                  }
                                : prev
                            )
                          }
                        >
                          {compareDraftB.headers.map((h: string) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium text-slate-700">عمود السعر</p>
                        <select
                          className="w-full rounded-xl border border-slate-300 p-3"
                          value={compareDraftB.mapping.price}
                          onChange={(e) =>
                            setCompareDraftB((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    mapping: { ...prev.mapping, price: e.target.value },
                                  }
                                : prev
                            )
                          }
                        >
                          {compareDraftB.headers.map((h: string) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium text-slate-700">عمود الخصم</p>
                        <select
                          className="w-full rounded-xl border border-slate-300 p-3"
                          value={compareDraftB.mapping.discount}
                          onChange={(e) =>
                            setCompareDraftB((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    mapping: { ...prev.mapping, discount: e.target.value },
                                  }
                                : prev
                            )
                          }
                        >
                          {compareDraftB.headers.map((h: string) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </div>

                      <button
                        className="w-full rounded-xl bg-slate-900 p-3 text-white hover:bg-slate-800"
                        onClick={() => finalizeCompareFile("B")}
                      >
                        اعتماد الملف
                      </button>
                    </>
                  )}
                </div>
              </Card>
            </div>

            <div className="flex flex-wrap justify-center gap-3">
              <button
                className="rounded-xl bg-slate-900 px-8 py-3 text-white disabled:opacity-50"
                onClick={runComparison}
                disabled={!compareA || !compareB}
              >
                ابدأ المقارنة
              </button>

              <button
                className="rounded-xl border border-slate-300 px-8 py-3 disabled:opacity-50"
                onClick={exportComparison}
                disabled={!compareResult}
              >
                تصدير CSV
              </button>
            </div>

            {compareResult && (
              <div className="space-y-6">
                <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-6">
                  <StatBox title="إجمالي ملف 1" value={compareResult.totalA} />
                  <StatBox title="إجمالي ملف 2" value={compareResult.totalB} />
                  <StatBox title="أصناف مشتركة" value={compareResult.sharedCount} />
                  <StatBox title="الأفضل في ملف 1" value={compareResult.aBetter} />
                  <StatBox title="الأفضل في ملف 2" value={compareResult.bBetter} />
                  <StatBox title="حد التشابه" value={`${compareResult.threshold}%`} />
                </div>

                <Card>
                  <div className="p-6">
                    <h3 className="text-2xl font-bold text-slate-900">نتائج المقارنة</h3>
                    <p className="mt-2 text-sm text-slate-500">
                      تختار لكل ملف عمود الاسم وعمود السعر وعمود الخصم، ثم تتم
                      المقارنة الذكية بنسبة تشابه 80% حتى مع اختلاف الكتابة.
                    </p>

                    <div className="mt-4 overflow-x-auto">
                      <table className="min-w-full border-collapse">
                        <thead>
                          <tr className="border-b border-slate-200 text-right text-sm text-slate-500">
                            <th className="p-3">صنف ملف 1</th>
                            <th className="p-3">صنف ملف 2</th>
                            <th className="p-3">التشابه</th>
                            <th className="p-3">خصم 1</th>
                            <th className="p-3">نهائي 1</th>
                            <th className="p-3">خصم 2</th>
                            <th className="p-3">نهائي 2</th>
                            <th className="p-3">الأفضل</th>
                          </tr>
                        </thead>
                        <tbody>
                          {compareResult.shared.map((row: any) => (
                            <tr key={row.key} className="border-b border-slate-100">
                              <td className="p-3 font-medium">{row.a.productName}</td>
                              <td className="p-3 font-medium">{row.b.productName}</td>
                              <td className="p-3">{row.similarity}%</td>
                              <td className="p-3">{row.a.discount}%</td>
                              <td className="p-3">{currency(row.a.finalPrice)} ج</td>
                              <td className="p-3">{row.b.discount}%</td>
                              <td className="p-3">{currency(row.b.finalPrice)} ج</td>
                              <td className="p-3">
                                {row.winner === "Equal" ? "متساوي" : row.winner}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </Card>
              </div>
            )}
          </div>
        )}

        {activeTab === "dashboard" && currentUser.role === "supplier" && (
          <div className="space-y-6">
            <SectionTitle title="لوحة المورد" desc="بيانات شركتك الفعالة خلال آخر 7 أيام." />
            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              <StatBox title="عدد الأصناف" value={myOffers.length} />
              <StatBox title="أفضل خصم" value={topOffer ? `${topOffer.discount}%` : "0%"} />
              <StatBox title="الخطة الحالية" value={myCompany?.plan || "Free"} />
              <StatBox title="الكود" value={myCompany?.code || "—"} />
            </div>
          </div>
        )}

        {activeTab === "admin" && currentUser.role === "admin" && (
          <div className="space-y-6">
            <SectionTitle title="لوحة الأدمن" desc="الأدمن هو المسؤول عن تكويد العملاء ومتابعة الملفات الفعالة." />
            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-5">
              <StatBox title="المستخدمون" value={users.length} />
              <StatBox title="الشركات الفعالة" value={validCompanies.length} />
              <StatBox title="العروض الفعالة" value={allOffers.length} />
              <StatBox title="المجموعات الذكية" value={productCount} />
              <StatBox title="تنبيهات" value={uploadReport?.invalid?.length || 0} hint="صفوف غير صالحة في آخر رفع" />
            </div>

            <Card>
              <div className="p-6">
                <h3 className="text-2xl font-bold text-slate-900">أفضل الشركات حسب متوسط الخصم</h3>
                <p className="mt-2 text-sm text-slate-500">
                  يعرض فقط العروض الفعالة خلال آخر 7 أيام.
                </p>

                <div className="mt-4 space-y-3">
                  {marketLeaderboard.map((company, index) => (
                    <div
                      key={company.id}
                      className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-slate-200 p-4"
                    >
                      <div>
                        <p className="font-semibold text-slate-900">
                          #{index + 1} {company.companyName}
                        </p>
                        <p className="text-sm text-slate-500">
                          الكود: {company.code || "—"} • {company.items} صنف • متوسط خصم{" "}
                          {company.avgDiscount}%
                        </p>
                      </div>
                      <div className="rounded-full bg-slate-100 px-4 py-1 text-sm">
                        {company.plan}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </Card>

            <Card>
              <div className="p-6">
                <h3 className="text-2xl font-bold text-slate-900">إدارة الشركات</h3>
                <p className="mt-2 text-sm text-slate-500">
                  رفع ملف جديد لنفس العميل أو نفس الكود يستبدل القديم تلقائيًا.
                </p>

                <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                  {companies.map((company) => (
                    <Card key={company.id}>
                      <div className="flex items-center justify-between gap-3 p-4">
                        <div>
                          <p className="font-semibold text-slate-900">{company.companyName}</p>
                          <p className="text-sm text-slate-500">
                            الكود: {company.code || "—"} • {company.offers.length} صنف •{" "}
                            {company.plan}
                          </p>
                        </div>
                        <button
                          className="rounded-lg border border-slate-300 px-3 py-2 hover:bg-slate-50"
                          onClick={() => removeCompany(company.id)}
                        >
                          حذف
                        </button>
                      </div>
                    </Card>
                  ))}
                </div>
              </div>
            </Card>
          </div>
        )}
      </div>
    </div>
  );
}
