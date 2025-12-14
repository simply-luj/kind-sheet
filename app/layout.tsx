import type { Metadata } from "next";
import { Jua, Noto_Sans_KR } from "next/font/google";
import "./globals.css";

const jua = Jua({
  weight: "400",
  subsets: ["latin"],
  variable: "--font-jua",
});

const notoSansKr = Noto_Sans_KR({
  subsets: ["latin"],
  variable: "--font-noto-sans-kr",
});

export const metadata: Metadata = {
  title: "Kind Sheet",
  description: "Excel processing made kind.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body
        className={`${jua.variable} ${notoSansKr.variable} font-sans bg-slate-900 text-slate-100 antialiased`}
      >
        {children}
      </body>
    </html>
  );
}
