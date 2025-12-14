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
  title: "Kind Sheet (카인드 시트) - 어려운 엑셀을 누구나 쉽게",
  description: "복잡한 엑셀 수식은 이제 그만. 나만의 버튼(Recipe) 하나로 엑셀 업무를 끝내보세요. 누구나 쉽게 사용하는 엑셀, 카인드 시트.",
  keywords: ["엑셀", "카인드시트", "Kind Sheet", "엑셀 자동화", "엑셀 수식", "스프레드시트"],
  icons: {
    icon: "/favicon.ico", // 파비콘이 있다면
  },
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
