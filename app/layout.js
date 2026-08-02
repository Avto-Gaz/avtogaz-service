export const metadata = {
  title: "Avtogaz Service",
  description: "Avtogaz Service — boshqaruv tizimi",
};

export const viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 1,
  themeColor: "#0B1220",
};

export default function RootLayout({ children }) {
  return (
    <html lang="uz">
      <body>{children}</body>
    </html>
  );
}
