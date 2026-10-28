import { NextResponse } from "next/server";
import { getSubmissions } from "@/lib/sec";

export async function GET(req: Request) {
  const cik = Number(new URL(req.url).searchParams.get("cik"));
  if (!cik) return NextResponse.json({ error: "Missing cik" }, { status: 400 });
  try {
    return NextResponse.json(await getSubmissions(cik));
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 502 });
  }
}
