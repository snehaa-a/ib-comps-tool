import { NextResponse } from "next/server";
import { suggestPeers } from "@/lib/peers";

export const maxDuration = 60;

export async function GET(req: Request) {
  const cik = Number(new URL(req.url).searchParams.get("cik"));
  if (!cik) return NextResponse.json({ error: "Missing cik" }, { status: 400 });
  try {
    return NextResponse.json(await suggestPeers(cik));
  } catch (e) {
    // Peer suggestions are a convenience; the student can still add peers by hand.
    return NextResponse.json({ sicDescription: "", peers: [], warning: (e as Error).message });
  }
}
