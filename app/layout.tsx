import type { Metadata } from 'next';
import { Geist_Mono, Noto_Sans_SC, Outfit } from 'next/font/google';
import './globals.css';

const outfit = Outfit({
  variable: '--font-outfit',
  subsets: ['latin'],
});

const notoSansSC = Noto_Sans_SC({
  variable: '--font-noto-sans-sc',
  subsets: ['latin'],
  preload: false,
});

const geistMono = Geist_Mono({
  variable: '--font-geist-mono',
  subsets: ['latin'],
});

export const metadata: Metadata = {
  title: '战斗设计知识库',
  description: '面向战斗策划的角色、怪物、玩法与游戏感权威资料研究台',
};

export const dynamic = 'force-static';

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body
        className={`${outfit.variable} ${notoSansSC.variable} ${geistMono.variable} antialiased`}
      >
        {children}
      </body>
    </html>
  );
}
