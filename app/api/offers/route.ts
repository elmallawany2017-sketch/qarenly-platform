import { NextResponse } from 'next/server';
import { listCompanies, replaceCompanyOffers } from '@/lib/store';
import { calcFinalPrice, normalizeName } from '@/lib/utils';

export async function GET() {
  return NextResponse.json({ companies: listCompanies() });
}

export async function POST(req: Request) {
  try {
    const body = await req.json();
    const offers = (body.rows || []).map((row: { productName: string; price: number; discount: number }) => ({
      productName: String(row.productName || '').trim(),
      normalizedName: normalizeName(row.productName),
      price: Number(row.price),
      discount: Number(row.discount),
      finalPrice: calcFinalPrice(Number(row.price), Number(row.discount)),
    }));
    const company = replaceCompanyOffers(body.companyName, body.email, offers);
    return NextResponse.json({ company });
  } catch {
    return NextResponse.json({ error: 'Failed to save offers' }, { status: 400 });
  }
}
