"use client";

import { useEffect, useState } from "react";

type Ticker = { cik: number; ticker: string; name: string };
type Company = Ticker & { sicDescription: string; fiscalYearEnd: string };
type Peer = Ticker & { reason: string };

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const fyLabel = (mmdd: string) => (mmdd.length === 4 ? `${MONTHS[Number(mmdd.slice(0, 2)) - 1]} ${Number(mmdd.slice(2))}` : "unknown");

async function getJson<T>(url: string): Promise<T> {
  const res = await fetch(url);
  const j = await res.json();
  if (!res.ok || j.error) throw new Error(j.error ?? `Request failed (${res.status})`);
  return j as T;
}

export default function Home() {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<Ticker[]>([]);
  const [company, setCompany] = useState<Company | null>(null);
  const [confirmed, setConfirmed] = useState(false);
  const [peers, setPeers] = useState<Peer[]>([]);
  const [peersLoading, setPeersLoading] = useState(false);
  const [newPeer, setNewPeer] = useState("");
  const [status, setStatus] = useState<"idle" | "working" | "done">("idle");
  const [error, setError] = useState("");

  useEffect(() => {
    if (!query.trim() || company) return setResults([]);
    const t = setTimeout(() => {
      getJson<{ results: Ticker[] }>(`/api/search?q=${encodeURIComponent(query)}`)
        .then((r) => setResults(r.results))
        .catch((e) => setError(e.message));
    }, 250);
    return () => clearTimeout(t);
  }, [query, company]);

  async function choose(t: Ticker) {
    setError("");
    setResults([]);
    try {
      const sub = await getJson<{ sicDescription: string; fiscalYearEnd: string }>(`/api/company?cik=${t.cik}`);
      setCompany({ ...t, sicDescription: sub.sicDescription, fiscalYearEnd: sub.fiscalYearEnd });
    } catch (e) {
      setError((e as Error).message);
    }
  }

  async function confirm() {
    if (!company) return;
    setConfirmed(true);
    setPeersLoading(true);
    try {
      const r = await getJson<{ peers: Peer[] }>(`/api/peers?cik=${company.cik}`);
      setPeers(r.peers);
    } catch {
      setPeers([]);
    } finally {
      setPeersLoading(false);
    }
  }

  async function addPeer() {
    const q = newPeer.trim();
    if (!q) return;
    setError("");
    try {
      const r = await getJson<{ results: Ticker[] }>(`/api/search?q=${encodeURIComponent(q)}`);
      const hit = r.results[0];
      if (!hit) throw new Error(`No SEC filer found for "${q}"`);
      if (hit.cik !== company?.cik && !peers.some((p) => p.cik === hit.cik)) {
        setPeers([...peers, { ...hit, reason: "added by you" }]);
      }
      setNewPeer("");
    } catch (e) {
      setError((e as Error).message);
    }
  }

  async function generate() {
    if (!company) return;
    setStatus("working");
    setError("");
    try {
      const res = await fetch("/api/workbook", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ target: company.ticker, peers: peers.map((p) => p.ticker) }),
      });
      if (!res.ok) throw new Error((await res.json()).error ?? "Could not build the workbook");
      const url = URL.createObjectURL(await res.blob());
      const a = document.createElement("a");
      a.href = url;
      a.download = `${company.ticker}_comps.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
      setStatus("done");
    } catch (e) {
      setError((e as Error).message);
      setStatus("idle");
    }
  }

  function reset() {
    setQuery("");
    setCompany(null);
    setConfirmed(false);
    setPeers([]);
    setStatus("idle");
    setError("");
  }

  return (
    <main>
      <h1>IB Comps Builder</h1>
      <p className="lead">Type a public company. Get a trading comps workbook built from its SEC filings, ready to evaluate.</p>

      <section className={`card ${confirmed ? "done" : ""}`}>
        <h2>1. Find the company</h2>
        {!company ? (
          <>
            <label htmlFor="q">Company name or ticker</label>
            <input id="q" type="text" value={query} placeholder="Nike or NKE" autoComplete="off" onChange={(e) => setQuery(e.target.value)} />
            <ul className="results">
              {results.map((r) => (
                <li key={r.cik}>
                  <button onClick={() => choose(r)}>
                    {r.name} · {r.ticker}
                  </button>
                </li>
              ))}
            </ul>
          </>
        ) : (
          <>
            <div className="match">
              <strong>
                {company.name} · {company.ticker}
              </strong>
              <span className="muted">
                Fiscal year ends {fyLabel(company.fiscalYearEnd)}
                {company.sicDescription ? ` · ${company.sicDescription}` : ""}
              </span>
            </div>
            {!confirmed ? (
              <div className="row">
                <button onClick={confirm}>Confirm</button>
                <button className="link" onClick={reset}>
                  Not this one
                </button>
              </div>
            ) : (
              <button className="link" onClick={reset}>
                Start over
              </button>
            )}
          </>
        )}
      </section>

      {confirmed && (
        <section className="card">
          <h2>2. Review the peers</h2>
          {peersLoading ? (
            <p className="muted">Finding companies in the same industry…</p>
          ) : (
            <>
              {!peers.length && <p className="muted">No suggestions found. Add peers by name or ticker below.</p>}
              <ul className="peers">
                {peers.map((p) => (
                  <li key={p.cik}>
                    <span>
                      {p.name} · {p.ticker}
                    </span>
                    <span className="muted">{p.reason}</span>
                    <button className="x" aria-label={`Remove ${p.ticker}`} onClick={() => setPeers(peers.filter((x) => x.cik !== p.cik))}>
                      ×
                    </button>
                  </li>
                ))}
              </ul>
              <div className="row">
                <input type="text" value={newPeer} placeholder="Add a company (name or ticker)" onChange={(e) => setNewPeer(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addPeer()} />
                <button onClick={addPeer}>Add</button>
              </div>
              <button onClick={generate} disabled={status === "working"}>
                Generate workbook
              </button>
            </>
          )}
        </section>
      )}

      {status !== "idle" && (
        <section className="card">
          <h2>{status === "working" ? "3. Generating" : "4. Download"}</h2>
          {status === "working" ? (
            <p className="muted">Pulling filings from SEC EDGAR, matching line items, and calculating LTM, EV and multiples. Usually under a minute.</p>
          ) : (
            <p>
              <span className="success">{company?.ticker}_comps.xlsx downloaded.</span>{" "}
              <span className="muted">Opens in Excel or Google Sheets. Check the Sources tab for every figure&apos;s filing.</span>
            </p>
          )}
        </section>
      )}

      {error && <p className="error">{error}</p>}

      <footer>Data from SEC EDGAR filings. For learning purposes; verify figures before relying on them.</footer>
    </main>
  );
}
