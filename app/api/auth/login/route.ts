import { NextResponse } from 'next/server';
import { findUser } from '@/lib/store';

export async function POST(req: Request) {
  const body = await req.json();
  const user = findUser(body.email, body.password);
  if (!user) {
    return NextResponse.json({ error: 'Invalid credentials' }, { status: 401 });
  }
  const { password, ...safeUser } = user;
  return NextResponse.json({ user: safeUser });
}
