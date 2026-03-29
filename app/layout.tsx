import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'Qarenly',
  description: 'Discount and price comparison platform for uploaded Excel catalogs.',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ar" dir="rtl">
      <body>{children}</body>
    </html>
  );
}
