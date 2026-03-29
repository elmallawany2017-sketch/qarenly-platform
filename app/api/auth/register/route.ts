import { NextResponse } from 'next/server';
import { createUser } from '@/lib/store';

export async function POST(req: Request) {
  try {
    const body = await req.json();
    const user = createUser(body);
    const { password, ...safeUser } = user;
    return NextResponse.json({ user: safeUser });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : 'Registration failed' }, { status: 400 });
  }
}
