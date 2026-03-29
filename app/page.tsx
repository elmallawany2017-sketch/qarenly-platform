'use client';

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

const guessColumn = (
  headers: string[],
  kind: "name" | "price" | "discount"
) => {
  const normalized = headers.map((h) => ({ raw: h, key: normalizeName(h) }));

  const patterns = {
    name: ["name", "product name", "اسم الصنف", "الصنف", "اسم المنتج"],
    price: ["price", "السعر", "سعر", "unit price", "buy price"],
    discount: ["discount", "الخصم", "discount %", "خصم", "نسبه الخصم"],
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

const seedCompanies: Company[] = [
  {
    id: crypto.randomUUID(),
    companyName: "Alpha Pharma",
    email: "alpha@qarenly.com",
    plan: "Pro",
    offers: [
      {
        productName: "Panadol Extra",
        normalizedName: normalizeName("Panadol Extra"),
        price: 100,
        discount: 20,
        finalPrice: 80,
      },
      {
        productName: "Augmentin 1g",
        normalizedName: normalizeName("Augmentin 1g"),
        price: 180,
        discount: 12,
        finalPrice: 158.4,
      },
      {
        productName: "Cetal 500",
        normalizedName: normalizeName("Cetal 500"),
        price: 30,
        discount: 5,
        finalPrice: 28.5,
      },
    ],
  },
  {
    id: crypto.randomUUID(),
    companyName: "Trust Med",
    email: "trust@qarenly.com",
    plan: "Business",
    offers: [
      {
        productName: "Panadol Exra",
        normalizedName: normalizeName("Panadol Exra"),
        price: 98,
        discount: 15,
        finalPrice: 83.3,
      },
      {
        productName: "Augmentin 1 gm",
        normalizedName: normalizeName("Augmentin 1 gm"),
        price: 175,
        discount: 8,
        finalPrice: 161,
      },
      {
        productName: "Brufen 400",
        normalizedName: normalizeName("Brufen 400"),
        price: 48,
        discount: 10,
        finalPrice: 43.2,
      },
    ],
  },
  {
    id: crypto.randomUUID(),
    companyName: "Market Plus",
    email: "market@qarenly.com",
    plan: "Free",
    offers: [
      {
        productName: "Panadol Extra",
        normalizedName: normalizeName("Panadol Extra"),
        price: 101,
        discount: 22,
        finalPrice: 78.78,
      },
      {
        productName: "Cetal500",
        normalizedName: normalizeName("Cetal500"),
        price: 31,
        discount: 8,
        finalPrice: 28.52,
      },
      {
        productName: "Brufen 400",
        normalizedName: normalizeName("Brufen 400"),
        price: 47,
        discount: 5,
        finalPrice: 44.65,
      },
    ],
  },
];

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
    <div className={`rounded-2xl border bg-white shadow-sm ${className}`}>
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
        <div className="mt-2 text-2xl font-bold">{value}</div>
        {hint ? <div className="mt-1 text-xs text-slate-400">{hint}</div> : null}
      </div>
    </Card>
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
  const [uploadReport, setUploadReport] = useState<any>(null);

  const [uploadDraft, setUploadDraft] = useState<any>(null);
  const [compareDraftA, setCompareDraftA] = useState<any>(null);
  const [compareDraftB, setCompareDraftB] = useState<any>(null);
  const [compareA, setCompareA] = useState<any>(null);
  const [compareB, setCompareB] = useState<any>(null);
  const [compareResult, setCompareResult] = useState<any>(null);

  const companyFileRef = useRef<HTMLInputElement>(null);
  const compareARef = useRef<HTMLInputElement>(null);
  const compareBRef = useRef<HTMLInputElement>(null);

  const allOffers = useMemo(() => {
    return companies.flatMap((company) =>
      company.offers.map((offer) => ({
        ...offer,
        companyName: company.companyName,
        companyId: company.id,
        plan: company.plan,
      }))
    );
  }, [companies]);

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

    const groups: any[] = [];

    for (const row of data) {
      let group = groups.find(
        (g) => similarityScore(g.anchor, row.productName) >= 0.8
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
      group.rows.sort((a: any, b: any) => {
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
  }, [allOffers, search, sortBy, filterCompany]);

  const productCount = groupedResults.length;

  const avgDiscount = allOffers.length
    ? (
        allOffers.reduce((sum, row) => sum + row.discount, 0) / allOffers.length
      ).toFixed(1)
    : "0";

  const myCompany = useMemo(() => {
    if (!currentUser || currentUser.role !== "supplier") return null;
    return companies.find(
      (c) => normalizeName(c.companyName) === normalizeName(currentUser.companyName)
    );
  }, [companies, currentUser]);

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
        (c) => normalizeName(c.companyName) === normalizeName(resolvedCompanyName)
      );

      if (existingIndex >= 0) {
        const clone = [...prev];
        clone[existingIndex] = { ...clone[existingIndex], offers: mappedOffers };
        return clone;
      }

      return [
        ...prev,
        {
          id: crypto.randomUUID(),
          companyName: resolvedCompanyName,
          email: currentUser?.email || "",
          plan: "Free",
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
    });

    setUploadDraft(null);
    setCompanyName("");
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
      label: which === "A" ? "ملف 1" : "ملف 2",
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
      name: draft.label,
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
        0.8
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

    if (rows.length) {
      downloadCsv(rows, "comparison-results.csv");
    }
  };

  const removeCompany = (id: string) => {
    setCompanies((prev) => prev.filter((c) => c.id !== id));
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-50 p-4 md:p-8">
        <div className="mx-auto grid max-w-7xl gap-6 lg:grid-cols-[1.1fr_.9fr]">
          <Card className="border-0 bg-gradient-to-br from-white to-slate-100 shadow-lg">
            <div className="p-8 md:p-10">
              <div className="mb-4 inline-block rounded-full bg-black px-4 py-1 text-sm text-white">
                Qarenly SaaS
              </div>
              <h1 className="text-3xl font-bold tracking-tight md:text-5xl">
                منصة ذكية لمقارنة الخصومات والأسعار
              </h1>
              <p className="mt-4 max-w-2xl text-base text-slate-600 md:text-lg">
                ارفع Excel وحدد بنفسك عمود الاسم وعمود السعر وعمود الخصم، ثم
                قارن الأصناف بنسبة تشابه 80% لمراعاة اختلاف الكتابة.
              </p>
              <div className="mt-6 grid gap-4 sm:grid-cols-3">
                <StatBox title="الشركات" value={companies.length} hint="موردون وعملاء" />
                <StatBox title="العروض" value={allOffers.length} hint="إجمالي العروض" />
                <StatBox title="متوسط الخصم" value={`${avgDiscount}%`} hint="على مستوى السوق" />
              </div>
            </div>
          </Card>

          <Card className="shadow-lg">
            <div className="p-6">
              <h2 className="text-2xl font-bold">
                {authMode === "login" ? "تسجيل الدخول" : "إنشاء حساب مورد"}
              </h2>
              <p className="mt-2 text-sm text-slate-500">
                جرب بحساب الأدمن أو أنشئ حساب مورد جديد.
              </p>

              <div className="mt-4 space-y-4">
                {authMode === "login" ? (
                  <>
                    <input
                      className="w-full rounded-xl border p-3"
                      placeholder="البريد الإلكتروني"
                      value={loginForm.email}
                      onChange={(e) =>
                        setLoginForm((p) => ({ ...p, email: e.target.value }))
                      }
                    />
                    <input
                      className="w-full rounded-xl border p-3"
                      type="password"
                      placeholder="كلمة المرور"
                      value={loginForm.password}
                      onChange={(e) =>
                        setLoginForm((p) => ({ ...p, password: e.target.value }))
                      }
                    />
                    <button
                      className="w-full rounded-xl bg-black p-3 text-white"
                      onClick={login}
                    >
                      دخول
                    </button>
                    <div className="rounded-xl border p-4 text-sm text-slate-500">
                      <p>حساب الأدمن: admin@qarenly.com</p>
                      <p>كلمة المرور: 123456</p>
                    </div>
                  </>
                ) : (
                  <>
                    <input
                      className="w-full rounded-xl border p-3"
                      placeholder="الاسم"
                      value={registerForm.name}
                      onChange={(e) =>
                        setRegisterForm((p) => ({ ...p, name: e.target.value }))
                      }
                    />
                    <input
                      className="w-full rounded-xl border p-3"
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
                      className="w-full rounded-xl border p-3"
                      placeholder="البريد الإلكتروني"
                      value={registerForm.email}
                      onChange={(e) =>
                        setRegisterForm((p) => ({ ...p, email: e.target.value }))
                      }
                    />
                    <input
                      className="w-full rounded-xl border p-3"
                      type="password"
                      placeholder="كلمة المرور"
                      value={registerForm.password}
                      onChange={(e) =>
                        setRegisterForm((p) => ({ ...p, password: e.target.value }))
                      }
                    />
                    <button
                      className="w-full rounded-xl bg-black p-3 text-white"
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
                  className="w-full rounded-xl border p-3"
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
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="mx-auto max-w-7xl space-y-8">
        <Card className="border-0 shadow-lg">
          <div className="flex flex-col gap-5 p-6 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="flex items-center gap-2">
                <div className="rounded-full bg-black px-4 py-1 text-sm text-white">
                  {currentUser.role === "admin" ? "Admin" : "Supplier"}
                </div>
                <div className="rounded-full bg-slate-100 px-4 py-1 text-sm">
                  {currentUser.companyName}
                </div>
              </div>
              <h1 className="mt-3 text-3xl font-bold">مرحبًا، {currentUser.name}</h1>
              <p className="mt-1 text-sm text-slate-500">
                رفع مرن للأعمدة + مقارنة بنسبة تشابه 80% بين الأصناف.
              </p>
            </div>
            <button className="rounded-xl border px-4 py-2" onClick={logout}>
              تسجيل خروج
            </button>
          </div>
        </Card>

        <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
          <button
            onClick={() => setActiveTab("market")}
            className={`rounded-xl p-3 ${activeTab === "market" ? "bg-black text-white" : "border bg-white"}`}
          >
            بحث السوق
          </button>
          <button
            onClick={() => setActiveTab("upload")}
            className={`rounded-xl p-3 ${activeTab === "upload" ? "bg-black text-white" : "border bg-white"}`}
          >
            رفع ملف
          </button>
          <button
            onClick={() => setActiveTab("compare")}
            className={`rounded-xl p-3 ${activeTab === "compare" ? "bg-black text-white" : "border bg-white"}`}
          >
            قارن ملفين
          </button>
          <button
            onClick={() => setActiveTab(currentUser.role === "admin" ? "admin" : "dashboard")}
            className={`rounded-xl p-3 ${activeTab === (currentUser.role === "admin" ? "admin" : "dashboard") ? "bg-black text-white" : "border bg-white"}`}
          >
            {currentUser.role === "admin" ? "الأدمن" : "لوحتي"}
          </button>
        </div>

        {activeTab === "market" && (
          <div className="space-y-6">
            <Card>
              <div className="p-6">
                <h2 className="text-2xl font-bold">محرك مقارنة السوق</h2>
                <p className="mt-2 text-sm text-slate-500">
                  عند البحث عن أي صنف ستظهر كل الشركات أو العملاء العارضين له، مع
                  اسم كل عميل وبجواره الخصم والسعر والسعر النهائي، مرتبين لسهولة
                  معرفة أعلى خصم.
                </p>

                <div className="mt-4 grid gap-4 lg:grid-cols-[1fr_220px_260px]">
                  <input
                    className="rounded-xl border p-3"
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="ابحث باسم الصنف أو الشركة..."
                  />

                  <select
                    className="rounded-xl border p-3"
                    value={sortBy}
                    onChange={(e) => setSortBy(e.target.value)}
                  >
                    <option value="discount">أعلى خصم</option>
                    <option value="final">أقل سعر نهائي</option>
                    <option value="name">اسم الصنف</option>
                  </select>

                  <select
                    className="rounded-xl border p-3"
                    value={filterCompany}
                    onChange={(e) => setFilterCompany(e.target.value)}
                  >
                    <option value="all">كل الشركات</option>
                    {companies.map((company) => (
                      <option key={company.id} value={company.id}>
                        {company.companyName}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </Card>

            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              <StatBox title="إجمالي الشركات" value={companies.length} />
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
                          <h3 className="text-xl font-bold">{group.displayName}</h3>
                          <p className="text-sm text-slate-500">
                            {group.rows.length} عروض متقاربة
                          </p>
                        </div>
                        <div className="rounded-full bg-slate-100 px-4 py-1 text-sm">
                          أفضل عرض:{" "}
                          {sortBy === "final"
                            ? `${currency(group.rows[0].finalPrice)} ج`
                            : `${group.rows[0].discount}%`}
                        </div>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="min-w-full border-collapse">
                          <thead>
                            <tr className="border-b text-right text-sm text-slate-500">
                              <th className="p-3">الترتيب</th>
                              <th className="p-3">الشركة / العميل</th>
                              <th className="p-3">اسم الصنف</th>
                              <th className="p-3">الخصم</th>
                              <th className="p-3">السعر</th>
                              <th className="p-3">السعر النهائي</th>
                            </tr>
                          </thead>
                          <tbody>
                            {group.rows.map((row: any, index: number) => (
                              <tr key={`${row.companyId}-${index}`} className="border-b">
                                <td className="p-3">{index + 1}</td>
                                <td className="p-3 font-medium">{row.companyName}</td>
                                <td className="p-3">{row.productName}</td>
                                <td className="p-3">{row.discount}%</td>
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
              <h2 className="text-2xl font-bold">رفع ملف شركة</h2>
              <p className="text-sm text-slate-500">
                بعد رفع الملف ستختار بنفسك: عمود الاسم + عمود السعر + عمود الخصم.
              </p>

              {currentUser.role !== "supplier" ? (
                <div className="grid gap-4 md:grid-cols-[1fr_auto]">
                  <input
                    className="rounded-xl border p-3"
                    value={companyName}
                    onChange={(e) => setCompanyName(e.target.value)}
                    placeholder="اسم الشركة أو العميل"
                  />
                  <button
                    className="rounded-xl bg-black px-4 py-3 text-white"
                    onClick={() => companyFileRef.current?.click()}
                  >
                    اختر ملف Excel
                  </button>
                </div>
              ) : (
                <div className="rounded-xl border p-4 text-sm text-slate-500">
                  سيتم ربط الملف تلقائيًا بشركتك: <strong>{currentUser.companyName}</strong>
                </div>
              )}

              {currentUser.role === "supplier" && (
                <button
                  className="rounded-xl bg-black px-4 py-3 text-white"
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
                <Card>
                  <div className="grid gap-4 p-5 md:grid-cols-3">
                    <div>
                      <p className="mb-2 text-sm font-medium">اختار عمود الاسم</p>
                      <select
                        className="w-full rounded-xl border p-3"
                        value={uploadDraft.mapping.name}
                        onChange={(e) =>
                          setUploadDraft((prev: any) => ({
                            ...prev,
                            mapping: { ...prev.mapping, name: e.target.value },
                          }))
                        }
                      >
                        {uploadDraft.headers.map((h: string) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <p className="mb-2 text-sm font-medium">اختار عمود السعر</p>
                      <select
                        className="w-full rounded-xl border p-3"
                        value={uploadDraft.mapping.price}
                        onChange={(e) =>
                          setUploadDraft((prev: any) => ({
                            ...prev,
                            mapping: { ...prev.mapping, price: e.target.value },
                          }))
                        }
                      >
                        {uploadDraft.headers.map((h: string) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <p className="mb-2 text-sm font-medium">اختار عمود الخصم</p>
                      <select
                        className="w-full rounded-xl border p-3"
                        value={uploadDraft.mapping.discount}
                        onChange={(e) =>
                          setUploadDraft((prev: any) => ({
                            ...prev,
                            mapping: { ...prev.mapping, discount: e.target.value },
                          }))
                        }
                      >
                        {uploadDraft.headers.map((h: string) => (
                          <option key={h} value={h}>
                            {h}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div className="flex gap-3 md:col-span-3">
                      <button
                        className="rounded-xl bg-black px-4 py-3 text-white"
                        onClick={finalizeUpload}
                      >
                        تأكيد الأعمدة ورفع البيانات
                      </button>
                      <button
                        className="rounded-xl border px-4 py-3"
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
                    <div className="grid gap-3 md:grid-cols-4">
                      <StatBox title="اسم الملف" value={uploadReport.fileName} />
                      <StatBox title="عمود الاسم" value={uploadReport.mapping.name} />
                      <StatBox title="عمود السعر" value={uploadReport.mapping.price} />
                      <StatBox title="عمود الخصم" value={uploadReport.mapping.discount} />
                    </div>
                    <div className="grid gap-3 md:grid-cols-2">
                      <StatBox title="صفوف صحيحة" value={uploadReport.valid} />
                      <StatBox title="صفوف بها أخطاء" value={uploadReport.invalid.length} />
                    </div>
                  </div>
                </Card>
              )}
            </div>
          </Card>
        )}

        {activeTab === "compare" && (
          <div className="space-y-6">
            <Card>
              <div className="p-5 text-sm text-slate-600">
                المنصة تطابق الأصناف بنسبة تشابه 80% حتى مع الاختلافات البسيطة في
                الكتابة.
              </div>
            </Card>

            <div className="grid gap-6 lg:grid-cols-2">
              <Card>
                <div className="p-6">
                  <h3 className="mb-4 text-xl font-bold">رفع الملف الأول</h3>
                  <button
                    className="w-full rounded-xl border p-3"
                    onClick={() => compareARef.current?.click()}
                  >
                    رفع الملف الأول
                  </button>
                </div>
              </Card>

              <Card>
                <div className="p-6">
                  <h3 className="mb-4 text-xl font-bold">رفع الملف الثاني</h3>
                  <button
                    className="w-full rounded-xl border p-3"
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
                  <h3 className="text-xl font-bold">إعداد الملف الأول</h3>
                  <p className="text-sm text-slate-500">
                    حدد عمود الاسم وعمود السعر وعمود الخصم قبل اعتماد الملف.
                  </p>

                  {!compareDraftA ? (
                    <p className="text-sm text-slate-500">لم يتم رفع الملف بعد.</p>
                  ) : (
                    <>
                      <div className="rounded-xl border p-4 text-sm text-slate-500">
                        {compareDraftA.fileName}
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium">عمود الاسم</p>
                        <select
                          className="w-full rounded-xl border p-3"
                          value={compareDraftA.mapping.name}
                          onChange={(e) =>
                            setCompareDraftA((prev: any) => ({
                              ...prev,
                              mapping: { ...prev.mapping, name: e.target.value },
                            }))
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
                        <p className="mb-2 text-sm font-medium">عمود السعر</p>
                        <select
                          className="w-full rounded-xl border p-3"
                          value={compareDraftA.mapping.price}
                          onChange={(e) =>
                            setCompareDraftA((prev: any) => ({
                              ...prev,
                              mapping: { ...prev.mapping, price: e.target.value },
                            }))
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
                        <p className="mb-2 text-sm font-medium">عمود الخصم</p>
                        <select
                          className="w-full rounded-xl border p-3"
                          value={compareDraftA.mapping.discount}
                          onChange={(e) =>
                            setCompareDraftA((prev: any) => ({
                              ...prev,
                              mapping: { ...prev.mapping, discount: e.target.value },
                            }))
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
                        className="w-full rounded-xl bg-black p-3 text-white"
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
                  <h3 className="text-xl font-bold">إعداد الملف الثاني</h3>
                  <p className="text-sm text-slate-500">
                    حدد عمود الاسم وعمود السعر وعمود الخصم قبل اعتماد الملف.
                  </p>

                  {!compareDraftB ? (
                    <p className="text-sm text-slate-500">لم يتم رفع الملف بعد.</p>
                  ) : (
                    <>
                      <div className="rounded-xl border p-4 text-sm text-slate-500">
                        {compareDraftB.fileName}
                      </div>

                      <div>
                        <p className="mb-2 text-sm font-medium">عمود الاسم</p>
                        <select
                          className="w-full rounded-xl border p-3"
                          value={compareDraftB.mapping.name}
                          onChange={(e) =>
                            setCompareDraftB((prev: any) => ({
                              ...prev,
                              mapping: { ...prev.mapping, name: e.target.value },
                            }))
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
                        <p className="mb-2 text-sm font-medium">عمود السعر</p>
                        <select
                          className="w-full rounded-xl border p-3"
                          value={compareDraftB.mapping.price}
                          onChange={(e) =>
                            setCompareDraftB((prev: any) => ({
                              ...prev,
                              mapping: { ...prev.mapping, price: e.target.value },
                            }))
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
                        <p className="mb-2 text-sm font-medium">عمود الخصم</p>
                        <select
                          className="w-full rounded-xl border p-3"
                          value={compareDraftB.mapping.discount}
                          onChange={(e) =>
                            setCompareDraftB((prev: any) => ({
                              ...prev,
                              mapping: { ...prev.mapping, discount: e.target.value },
                            }))
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
                        className="w-full rounded-xl bg-black p-3 text-white"
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
                className="rounded-xl bg-black px-8 py-3 text-white disabled:opacity-50"
                onClick={runComparison}
                disabled={!compareA || !compareB}
              >
                ابدأ المقارنة
              </button>

              <button
                className="rounded-xl border px-8 py-3 disabled:opacity-50"
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
                    <h3 className="text-2xl font-bold">نتائج المقارنة</h3>
                    <p className="mt-2 text-sm text-slate-500">
                      تختار لكل ملف عمود الاسم وعمود السعر وعمود الخصم، ثم تتم
                      المقارنة الذكية بنسبة تشابه 80% حتى مع اختلاف الكتابة.
                    </p>

                    <div className="mt-4 overflow-x-auto">
                      <table className="min-w-full border-collapse">
                        <thead>
                          <tr className="border-b text-right text-sm text-slate-500">
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
                            <tr key={row.key} className="border-b">
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
            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              <StatBox title="عدد الأصناف" value={myOffers.length} />
              <StatBox title="أفضل خصم" value={topOffer ? `${topOffer.discount}%` : "0%"} />
              <StatBox title="الخطة الحالية" value={myCompany?.plan || "Free"} />
              <StatBox title="اسم الشركة" value={myCompany?.companyName || currentUser.companyName} />
            </div>
          </div>
        )}

        {activeTab === "admin" && currentUser.role === "admin" && (
          <div className="space-y-6">
            <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-5">
              <StatBox title="المستخدمون" value={users.length} />
              <StatBox title="الشركات" value={companies.length} />
              <StatBox title="العروض" value={allOffers.length} />
              <StatBox title="المجموعات الذكية" value={productCount} />
              <StatBox
                title="تنبيهات"
                value={uploadReport?.invalid?.length || 0}
                hint="صفوف غير صالحة في آخر رفع"
              />
            </div>

            <Card>
              <div className="p-6">
                <h3 className="text-2xl font-bold">إدارة الشركات</h3>
                <p className="mt-2 text-sm text-slate-500">
                  حذف أو متابعة كل شركة مضافة داخل المنصة.
                </p>

                <div className="mt-4 grid gap-3 md:grid-cols-2 xl:grid-cols-3">
                  {companies.map((company) => (
                    <Card key={company.id}>
                      <div className="flex items-center justify-between gap-3 p-4">
                        <div>
                          <p className="font-semibold">{company.companyName}</p>
                          <p className="text-sm text-slate-500">
                            {company.offers.length} صنف • {company.plan}
                          </p>
                        </div>
                        <button
                          className="rounded-lg border px-3 py-2"
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
