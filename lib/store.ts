import { calcFinalPrice, normalizeName } from './utils';
import type { AppUser, Company, Offer } from './types';

declare global {
  var qarenlyStore:
    | {
        users: AppUser[];
        companies: Company[];
      }
    | undefined;
}

const seedCompanies: Company[] = [
  {
    id: 'c1',
    companyName: 'Alpha Pharma',
    email: 'alpha@qarenly.com',
    plan: 'Pro',
    offers: [
      { productName: 'Panadol Extra', normalizedName: normalizeName('Panadol Extra'), price: 100, discount: 20, finalPrice: calcFinalPrice(100, 20) },
      { productName: 'Augmentin 1g', normalizedName: normalizeName('Augmentin 1g'), price: 180, discount: 12, finalPrice: calcFinalPrice(180, 12) },
      { productName: 'Cetal 500', normalizedName: normalizeName('Cetal 500'), price: 30, discount: 5, finalPrice: calcFinalPrice(30, 5) },
    ],
  },
  {
    id: 'c2',
    companyName: 'Trust Med',
    email: 'trust@qarenly.com',
    plan: 'Business',
    offers: [
      { productName: 'Panadol Extra', normalizedName: normalizeName('Panadol Extra'), price: 98, discount: 15, finalPrice: calcFinalPrice(98, 15) },
      { productName: 'Augmentin 1g', normalizedName: normalizeName('Augmentin 1g'), price: 175, discount: 8, finalPrice: calcFinalPrice(175, 8) },
      { productName: 'Brufen 400', normalizedName: normalizeName('Brufen 400'), price: 48, discount: 10, finalPrice: calcFinalPrice(48, 10) },
    ],
  },
  {
    id: 'c3',
    companyName: 'Market Plus',
    email: 'market@qarenly.com',
    plan: 'Free',
    offers: [
      { productName: 'Panadol Extra', normalizedName: normalizeName('Panadol Extra'), price: 101, discount: 22, finalPrice: calcFinalPrice(101, 22) },
      { productName: 'Cetal 500', normalizedName: normalizeName('Cetal 500'), price: 31, discount: 8, finalPrice: calcFinalPrice(31, 8) },
      { productName: 'Brufen 400', normalizedName: normalizeName('Brufen 400'), price: 47, discount: 5, finalPrice: calcFinalPrice(47, 5) },
    ],
  },
];

const seedUsers: AppUser[] = [
  { id: 'u1', name: 'Admin', email: 'admin@qarenly.com', password: '123456', role: 'admin', companyName: 'Qarenly' },
  { id: 'u2', name: 'Alpha Pharma', email: 'alpha@qarenly.com', password: '123456', role: 'supplier', companyName: 'Alpha Pharma' },
  { id: 'u3', name: 'Trust Med', email: 'trust@qarenly.com', password: '123456', role: 'supplier', companyName: 'Trust Med' },
];

if (!global.qarenlyStore) {
  global.qarenlyStore = { users: seedUsers, companies: seedCompanies };
}

export const store = global.qarenlyStore;

export function listCompanies() {
  return store.companies;
}

export function listUsers() {
  return store.users.map(({ password, ...rest }) => rest);
}

export function findUser(email: string, password?: string) {
  return store.users.find((u) => u.email === email && (password ? u.password === password : true));
}

export function createUser(input: Omit<AppUser, 'id' | 'role'>) {
  const exists = store.users.some((u) => u.email === input.email);
  if (exists) throw new Error('Email already exists');
  const user: AppUser = {
    id: `u_${Date.now()}`,
    role: 'supplier',
    ...input,
  };
  store.users.push(user);
  const companyExists = store.companies.some((c) => normalizeName(c.companyName) === normalizeName(input.companyName));
  if (!companyExists) {
    store.companies.push({
      id: `c_${Date.now()}`,
      companyName: input.companyName,
      email: input.email,
      plan: 'Free',
      offers: [],
    });
  }
  return user;
}

export function replaceCompanyOffers(companyName: string, email: string, offers: Offer[]) {
  const index = store.companies.findIndex((c) => normalizeName(c.companyName) === normalizeName(companyName));
  if (index >= 0) {
    store.companies[index] = { ...store.companies[index], offers, email };
    return store.companies[index];
  }
  const company: Company = { id: `c_${Date.now()}`, companyName, email, plan: 'Free', offers };
  store.companies.push(company);
  return company;
}
