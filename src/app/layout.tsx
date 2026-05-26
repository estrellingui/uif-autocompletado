import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "UIF Autocompletado | Planilla desde PDF",
  description:
    "Completá automáticamente la planilla UIF a partir de boletos de compraventa o escrituras en PDF.",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="es">
      <body className="min-h-screen bg-gray-50">{children}</body>
    </html>
  );
}
