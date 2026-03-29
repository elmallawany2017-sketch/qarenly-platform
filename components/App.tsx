'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { formatMoney, normalizeName } from '@/lib/utils';

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
  plan: 'Free' | 'Pro' | 'Business';
  offers: Offer[];
};

type SafeUser = {
  id: string;
  name: string;
  email: string;
  role: 'admin' | 'supplier';
  companyName: string;
};

type CompareResult = {
  shared: Array<{ productName: string; a: Offer; b: Offer; winner: string }>;
  onlyA: Offer[];
  onlyB: Offer[];
  aBetter: number;
  bBetter: number;
  equal: number;
  totalA: number;
  totalB: number;
};

async function parseExcelFile(file: File) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' }) as Record<string, unknown>[];
  return rows.map((row) => {
    const lowered = Object.keys(row).reduce<Record<string, unknown>>((acc, key) => {
      acc[key.toLowerCase().trim()] = row[key];
      return acc;
    }, {});
    const productName = String(lowered.name ?? lowered['اسم الصنف'] ?? lowered['product name'] ?? lowered['الصنف'] ?? '').trim();
    const price = Number(lowered.price ?? lowered['السعر'] ?? 0);
    const discount = Number(lowered.discount ?? lowered['الخصم'] ?? 0);
    return { productName, price, discount };
  }).filter((row) => row.productName && row.price > 0 && row.discount >= 0 && row.discount <= 100);
}

export default function App() {
  const [companies, setCompanies] = useState<Company[]>([]);
  const [currentUser, setCurrentUser] = useState<SafeUser | null>(null);
  const [authMode, setAuthMode] = useState<'login' | 'register'>('login');
  const [activeTab, setActiveTab] = useState<'market' | 'upload' | 'compare' | 'dashboard' | 'admin'>('market');
  const [loginForm, setLoginForm] = useState({ email: 'admin@qarenly.com', password: '123456' });
  const [registerForm, setRegisterForm] = useState({ name: '', companyName: '', email: '', password: '' });
  const [authError, setAuthError] = useState('');
  const [search, setSearch] = useState('');
  const [sortBy, setSortBy] = useState<'discount' | 'final'>('discount');
  const [filterCompany, setFilterCompany] = useState('all');
  const [companyName, setCompanyName] = useState('');
  const [uploadMessage, setUploadMessage] = useState('');
  const [compareA, setCompareA] = useState<{ name: string; rows: { productName: string; price: number; discount: number }[] } | null>(null);
  const [compareB, setCompareB] = useState<{ name: string; rows: { productName: string; price: number; discount: number }[] } | null>(null);
  const [compareResult, setCompareResult] = useState<CompareResult | null>(null);
  const [loading, setLoading] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const compareARef = useRef<HTMLInputElement>(null);
  const compareBRef = useRef<HTMLInputElement>(null);

  async function refreshCompanies() {
    const response = await fetch('/api/offers');
    const data = await response.json();
    setCompanies(data.companies);
  }

  useEffect(() => {
    refreshCompanies();
  }, []);

  const allOffers = useMemo(() => companies.flatMap((company) => company.offers.map((offer) => ({ ...offer, companyName: company.companyName, companyId: company.id, plan: company.plan }))), [companies]);
  const avgDiscount = useMemo(() => allOffers.length ? (allOffers.reduce((s, r) => s + r.discount, 0) / allOffers.length).toFixed(1) : '0', [allOffers]);
  const productCount = useMemo(() => new Set(allOffers.map((o) => o.normalizedName)).size, [allOffers]);
  const myCompany = useMemo(() => currentUser?.role === 'supplier' ? companies.find((c) => normalizeName(c.companyName) === normalizeName(currentUser.companyName)) : null, [companies, currentUser]);

  const groupedResults = useMemo(() => {
    let rows = allOffers;
    const query = normalizeName(search);
    if (query) {
      rows = rows.filter((row) => row.normalizedName.includes(query) || normalizeName(row.companyName).includes(query));
    }
    if (filterCompany !== 'all') {
      rows = rows.filter((row) => row.companyId === filterCompany);
    }
    rows = [...rows].sort((a, b) => sortBy === 'discount' ? b.discount - a.discount || a.finalPrice - b.finalPrice : a.finalPrice - b.finalPrice || b.discount - a.discount);

    const map = new Map<string, typeof rows>();
    for (const row of rows) {
      if (!map.has(row.normalizedName)) map.set(row.normalizedName, []);
      map.get(row.normalizedName)!.push(row);
    }
    return [...map.values()].map((rows) => ({ key: rows[0].normalizedName, displayName: rows[0].productName, rows }));
  }, [allOffers, search, sortBy, filterCompany]);

  async function login() {
    setLoading(true);
    setAuthError('');
    const response = await fetch('/api/auth/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(loginForm) });
    const data = await response.json();
    setLoading(false);
    if (!response.ok) {
      setAuthError(data.error || 'Login failed');
      return;
    }
    setCurrentUser(data.user);
    setActiveTab(data.user.role === 'admin' ? 'admin' : 'dashboard');
  }

  async function register() {
    setLoading(true);
    setAuthError('');
    const response = await fetch('/api/auth/register', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(registerForm) });
    const data = await response.json();
    setLoading(false);
    if (!response.ok) {
      setAuthError(data.error || 'Registration failed');
      return;
    }
    setCurrentUser(data.user);
    setActiveTab('dashboard');
    await refreshCompanies();
  }

  async function uploadCompanyFile(file: File) {
    const rows = await parseExcelFile(file);
    const resolvedCompanyName = currentUser?.role === 'supplier' ? currentUser.companyName : (companyName.trim() || file.name.replace(/\.[^.]+$/, ''));
    const email = currentUser?.email || 'manual@qarenly.com';
    const response = await fetch('/api/offers', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ companyName: resolvedCompanyName, email, rows }),
    });
    const data = await response.json();
    if (!response.ok) {
      setUploadMessage('حدث خطأ أثناء حفظ الملف.');
      return;
    }
    setUploadMessage(`تم رفع ${rows.length} صنف للشركة ${data.company.companyName}`);
    await refreshCompanies();
  }

  async function runCompare() {
    if (!compareA || !compareB) return;
    const response = await fetch('/api/compare', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileA: compareA.rows, fileB: compareB.rows }),
    });
    const data = await response.json();
    setCompareResult(data);
  }

  const marketLeaderboard = useMemo(() => {
    return companies
      .map((company) => {
        const avg = company.offers.length ? company.offers.reduce((sum, row) => sum + row.discount, 0) / company.offers.length : 0;
        return { ...company, avgDiscount: avg.toFixed(1) };
      })
      .sort((a, b) => Number(b.avgDiscount) - Number(a.avgDiscount));
  }, [companies]);

  return (
    <main className="container">
      {!currentUser ? (
        <div className="grid-2">
          <section className="card hero card-pad">
            <span className="badge">Qarenly Full Stack</span>
            <h1 className="title mt-16">منصة مقارنة الأسعار والخصومات برفع Excel</h1>
            <p className="subtle">جاهزة كمشروع Next.js كامل بواجهات، API routes، تسجيل دخول، رفع عروض، مقارنة ملفين، ولوحة أدمن.</p>
            <div className="grid-3 mt-24">
              <div className="stat"><h3>الشركات</h3><p>{companies.length}</p></div>
              <div className="stat"><h3>العروض</h3><p>{allOffers.length}</p></div>
              <div className="stat"><h3>متوسط الخصم</h3><p>{avgDiscount}%</p></div>
            </div>
          </section>
          <section className="card card-pad">
            <h2 style={{ marginTop: 0 }}>{authMode === 'login' ? 'تسجيل الدخول' : 'إنشاء حساب مورد'}</h2>
            <p className="subtle small">حساب الأدمن التجريبي: admin@qarenly.com / 123456</p>
            {authMode === 'login' ? (
              <div className="grid mt-16">
                <input className="input" placeholder="البريد الإلكتروني" value={loginForm.email} onChange={(e) => setLoginForm((p) => ({ ...p, email: e.target.value }))} />
                <input className="input" type="password" placeholder="كلمة المرور" value={loginForm.password} onChange={(e) => setLoginForm((p) => ({ ...p, password: e.target.value }))} />
                <button className="button" onClick={login} disabled={loading}>دخول</button>
              </div>
            ) : (
              <div className="grid mt-16">
                <input className="input" placeholder="الاسم" value={registerForm.name} onChange={(e) => setRegisterForm((p) => ({ ...p, name: e.target.value }))} />
                <input className="input" placeholder="اسم الشركة" value={registerForm.companyName} onChange={(e) => setRegisterForm((p) => ({ ...p, companyName: e.target.value }))} />
                <input className="input" placeholder="البريد الإلكتروني" value={registerForm.email} onChange={(e) => setRegisterForm((p) => ({ ...p, email: e.target.value }))} />
                <input className="input" type="password" placeholder="كلمة المرور" value={registerForm.password} onChange={(e) => setRegisterForm((p) => ({ ...p, password: e.target.value }))} />
                <button className="button" onClick={register} disabled={loading}>إنشاء الحساب</button>
              </div>
            )}
            {authError ? <div className="notice mt-16">{authError}</div> : null}
            <button className="button secondary mt-16" onClick={() => { setAuthMode(authMode === 'login' ? 'register' : 'login'); setAuthError(''); }}>
              {authMode === 'login' ? 'إنشاء حساب جديد' : 'لديك حساب؟ سجل دخول'}
            </button>
          </section>
        </div>
      ) : (
        <>
          <section className="card card-pad">
            <div className="between">
              <div>
                <div className="row">
                  <span className="badge">{currentUser.role === 'admin' ? 'Admin' : 'Supplier'}</span>
                  <span className="badge">{currentUser.companyName}</span>
                </div>
                <h1 className="title mt-16">مرحبًا، {currentUser.name}</h1>
                <p className="subtle">لوحة تشغيل كاملة للمنصة مع السوق، الرفع، المقارنة، والإدارة.</p>
              </div>
              <button className="button secondary" onClick={() => { setCurrentUser(null); setActiveTab('market'); }}>تسجيل خروج</button>
            </div>
          </section>

          <div className="tabs mt-24">
            {['market', 'upload', 'compare', currentUser.role === 'admin' ? 'admin' : 'dashboard'].map((tab) => (
              <button key={tab} className={`tab ${activeTab === tab ? 'active' : ''}`} onClick={() => setActiveTab(tab as typeof activeTab)}>
                {tab === 'market' ? 'بحث السوق' : tab === 'upload' ? 'رفع ملف' : tab === 'compare' ? 'قارن ملفين' : tab === 'admin' ? 'الأدمن' : 'لوحتي'}
              </button>
            ))}
          </div>

          {activeTab === 'market' && (
            <section className="grid mt-24">
              <div className="card card-pad">
                <div className="between">
                  <div>
                    <h2 style={{ margin: 0 }}>محرك السوق</h2>
                    <p className="subtle small">ابحث عن أي صنف وشاهد العروض مرتبة.</p>
                  </div>
                </div>
                <div className="grid-3 mt-16">
                  <input className="input" placeholder="ابحث بالصنف أو الشركة" value={search} onChange={(e) => setSearch(e.target.value)} />
                  <select className="select" value={sortBy} onChange={(e) => setSortBy(e.target.value as 'discount' | 'final')}>
                    <option value="discount">أعلى خصم</option>
                    <option value="final">أقل سعر نهائي</option>
                  </select>
                  <select className="select" value={filterCompany} onChange={(e) => setFilterCompany(e.target.value)}>
                    <option value="all">كل الشركات</option>
                    {companies.map((company) => <option key={company.id} value={company.id}>{company.companyName}</option>)}
                  </select>
                </div>
              </div>

              <div className="grid-4">
                <div className="stat"><h3>الشركات</h3><p>{companies.length}</p></div>
                <div className="stat"><h3>العروض</h3><p>{allOffers.length}</p></div>
                <div className="stat"><h3>أصناف فريدة</h3><p>{productCount}</p></div>
                <div className="stat"><h3>متوسط الخصم</h3><p>{avgDiscount}%</p></div>
              </div>

              {groupedResults.map((group) => (
                <div className="card card-pad" key={group.key}>
                  <div className="between">
                    <div>
                      <h3 style={{ margin: 0 }}>{group.displayName}</h3>
                      <p className="subtle small">{group.rows.length} عروض</p>
                    </div>
                    <span className="badge">أفضل عرض: {sortBy === 'final' ? `${formatMoney(group.rows[0].finalPrice)} ج` : `${group.rows[0].discount}%`}</span>
                  </div>
                  <div className="table-wrap mt-16">
                    <table>
                      <thead>
                        <tr><th>الترتيب</th><th>الشركة</th><th>السعر</th><th>الخصم</th><th>السعر النهائي</th><th>الخطة</th></tr>
                      </thead>
                      <tbody>
                        {group.rows.map((row, index) => (
                          <tr key={`${row.companyId}-${index}`}>
                            <td>{index + 1}</td>
                            <td>{row.companyName}</td>
                            <td>{formatMoney(row.price)} ج</td>
                            <td>{row.discount}%</td>
                            <td>{formatMoney(row.finalPrice)} ج</td>
                            <td>{row.plan}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))}
            </section>
          )}

          {activeTab === 'upload' && (
            <section className="card card-pad mt-24">
              <h2 style={{ marginTop: 0 }}>رفع ملف Excel</h2>
              <p className="subtle small">الأعمدة المطلوبة: name, price, discount</p>
              <div className="grid-2 mt-16">
                {currentUser.role === 'admin' ? (
                  <input className="input" placeholder="اسم الشركة" value={companyName} onChange={(e) => setCompanyName(e.target.value)} />
                ) : (
                  <div className="notice">سيتم ربط الملف تلقائيًا بشركتك: {currentUser.companyName}</div>
                )}
                <button className="button" onClick={() => fileRef.current?.click()}>اختر الملف</button>
              </div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={(e) => e.target.files?.[0] && uploadCompanyFile(e.target.files[0])} />
              {uploadMessage ? <div className="notice mt-16">{uploadMessage}</div> : null}
            </section>
          )}

          {activeTab === 'compare' && (
            <section className="grid mt-24">
              <div className="grid-2">
                <div className="card card-pad">
                  <h3 style={{ marginTop: 0 }}>الملف الأول</h3>
                  <button className="button secondary" onClick={() => compareARef.current?.click()}>رفع الملف الأول</button>
                  <input ref={compareARef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={async (e) => {
                    if (!e.target.files?.[0]) return;
                    const rows = await parseExcelFile(e.target.files[0]);
                    setCompareA({ name: 'File 1', rows });
                  }} />
                  {compareA ? <p className="subtle mt-16">{compareA.rows.length} صف صحيح</p> : null}
                </div>
                <div className="card card-pad">
                  <h3 style={{ marginTop: 0 }}>الملف الثاني</h3>
                  <button className="button secondary" onClick={() => compareBRef.current?.click()}>رفع الملف الثاني</button>
                  <input ref={compareBRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={async (e) => {
                    if (!e.target.files?.[0]) return;
                    const rows = await parseExcelFile(e.target.files[0]);
                    setCompareB({ name: 'File 2', rows });
                  }} />
                  {compareB ? <p className="subtle mt-16">{compareB.rows.length} صف صحيح</p> : null}
                </div>
              </div>
              <button className="button" onClick={runCompare} disabled={!compareA || !compareB}>ابدأ المقارنة</button>
              {compareResult ? (
                <>
                  <div className="grid-5">
                    <div className="stat"><h3>إجمالي ملف 1</h3><p>{compareResult.totalA}</p></div>
                    <div className="stat"><h3>إجمالي ملف 2</h3><p>{compareResult.totalB}</p></div>
                    <div className="stat"><h3>مشتركة</h3><p>{compareResult.shared.length}</p></div>
                    <div className="stat"><h3>الأفضل في 1</h3><p>{compareResult.aBetter}</p></div>
                    <div className="stat"><h3>الأفضل في 2</h3><p>{compareResult.bBetter}</p></div>
                  </div>
                  <div className="card card-pad">
                    <div className="table-wrap">
                      <table>
                        <thead>
                          <tr><th>الصنف</th><th>سعر 1</th><th>خصم 1</th><th>نهائي 1</th><th>سعر 2</th><th>خصم 2</th><th>نهائي 2</th><th>الأفضل</th></tr>
                        </thead>
                        <tbody>
                          {compareResult.shared.map((row, idx) => (
                            <tr key={idx}>
                              <td>{row.productName}</td>
                              <td>{formatMoney(row.a.price)} ج</td>
                              <td>{row.a.discount}%</td>
                              <td>{formatMoney(row.a.finalPrice)} ج</td>
                              <td>{formatMoney(row.b.price)} ج</td>
                              <td>{row.b.discount}%</td>
                              <td>{formatMoney(row.b.finalPrice)} ج</td>
                              <td>{row.winner}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              ) : null}
            </section>
          )}

          {activeTab === 'dashboard' && currentUser.role === 'supplier' && (
            <section className="grid mt-24">
              <div className="grid-4">
                <div className="stat"><h3>اسم الشركة</h3><p style={{ fontSize: 20 }}>{myCompany?.companyName || currentUser.companyName}</p></div>
                <div className="stat"><h3>عدد الأصناف</h3><p>{myCompany?.offers.length || 0}</p></div>
                <div className="stat"><h3>الخطة الحالية</h3><p>{myCompany?.plan || 'Free'}</p></div>
                <div className="stat"><h3>أفضل خصم</h3><p>{myCompany?.offers.length ? `${Math.max(...myCompany.offers.map((o) => o.discount))}%` : '0%'}</p></div>
              </div>
              <div className="card card-pad">
                <h2 style={{ marginTop: 0 }}>أصناف شركتك</h2>
                <div className="table-wrap mt-16">
                  <table>
                    <thead><tr><th>الصنف</th><th>السعر</th><th>الخصم</th><th>السعر النهائي</th></tr></thead>
                    <tbody>
                      {(myCompany?.offers || []).map((row, idx) => (
                        <tr key={idx}><td>{row.productName}</td><td>{formatMoney(row.price)} ج</td><td>{row.discount}%</td><td>{formatMoney(row.finalPrice)} ج</td></tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          )}

          {activeTab === 'admin' && currentUser.role === 'admin' && (
            <section className="grid mt-24">
              <div className="grid-4">
                <div className="stat"><h3>المستخدمون</h3><p>{companies.length + 1}</p></div>
                <div className="stat"><h3>الشركات</h3><p>{companies.length}</p></div>
                <div className="stat"><h3>العروض</h3><p>{allOffers.length}</p></div>
                <div className="stat"><h3>أصناف فريدة</h3><p>{productCount}</p></div>
              </div>
              <div className="card card-pad">
                <h2 style={{ marginTop: 0 }}>ترتيب الشركات حسب متوسط الخصم</h2>
                <div className="table-wrap mt-16">
                  <table>
                    <thead><tr><th>الترتيب</th><th>الشركة</th><th>الخطة</th><th>عدد الأصناف</th><th>متوسط الخصم</th></tr></thead>
                    <tbody>
                      {marketLeaderboard.map((company, idx) => (
                        <tr key={company.id}><td>{idx + 1}</td><td>{company.companyName}</td><td>{company.plan}</td><td>{company.offers.length}</td><td>{company.avgDiscount}%</td></tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </section>
          )}
        </>
      )}
    </main>
  );
}
