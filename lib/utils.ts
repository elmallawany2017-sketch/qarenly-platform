export const normalizeName = (value: string) =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\u0640]/g, '')
    .replace(/[أإآ]/g, 'ا')
    .replace(/[ة]/g, 'ه')
    .replace(/[ى]/g, 'ي')
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();

export const calcFinalPrice = (price: number, discount: number) => {
  return +(price - price * (discount / 100)).toFixed(2);
};

export const formatMoney = (value: number) =>
  new Intl.NumberFormat('en-EG', { maximumFractionDigits: 2 }).format(value);
