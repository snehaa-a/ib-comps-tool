import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "IB Comps Builder",
  description: "Turn a ticker into a trading comps workbook built from SEC filings.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
