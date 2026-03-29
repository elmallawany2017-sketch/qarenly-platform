export type Offer = {
  productName: string;
  normalizedName: string;
  price: number;
  discount: number;
  finalPrice: number;
};

export type Company = {
  id: string;
  companyName: string;
  email: string;
  plan: 'Free' | 'Pro' | 'Business';
  offers: Offer[];
};

export type AppUser = {
  id: string;
  name: string;
  email: string;
  password: string;
  role: 'admin' | 'supplier';
  companyName: string;
};
