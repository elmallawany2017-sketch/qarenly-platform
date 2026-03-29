import { NextResponse } from 'next/server';
import { calcFinalPrice, normalizeName } from '@/lib/utils';

type Row = { productName: string; price: number; discount: number };

function normalizeRows(rows: Row[]) {
  return rows.map((row) => ({
    productName: String(row.productName || '').trim(),
    normalizedName: normalizeName(row.productName),
    price: Number(row.price),
    discount: Number(row.discount),
    finalPrice: calcFinalPrice(Number(row.price), Number(row.discount)),
  }));
}

export async function POST(req: Request) {
  const body = await req.json();
  const aRows = normalizeRows(body.fileA || []);
  const bRows = normalizeRows(body.fileB || []);

  const aMap = new Map(aRows.map((r) => [r.normalizedName, r]));
  const bMap = new Map(bRows.map((r) => [r.normalizedName, r]));
  const keys = new Set([...aMap.keys(), ...bMap.keys()]);

  const shared = [] as Array<Record<string, unknown>>;
  const onlyA = [] as typeof aRows;
  const onlyB = [] as typeof bRows;
  let aBetter = 0;
  let bBetter = 0;
  let equal = 0;

  for (const key of keys) {
    const a = aMap.get(key);
    const b = bMap.get(key);
    if (a && b) {
      let winner = 'Equal';
      if (a.discount > b.discount) winner = 'File 1';
      else if (b.discount > a.discount) winner = 'File 2';
      else if (a.finalPrice < b.finalPrice) winner = 'File 1';
      else if (b.finalPrice < a.finalPrice) winner = 'File 2';
      if (winner === 'File 1') aBetter += 1;
      else if (winner === 'File 2') bBetter += 1;
      else equal += 1;
      shared.push({ productName: a.productName || b.productName, a, b, winner });
    } else if (a) {
      onlyA.push(a);
    } else if (b) {
      onlyB.push(b);
    }
  }

  return NextResponse.json({ shared, onlyA, onlyB, aBetter, bBetter, equal, totalA: aRows.length, totalB: bRows.length });
}
