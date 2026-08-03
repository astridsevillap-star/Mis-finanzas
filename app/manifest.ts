import type { MetadataRoute } from "next";

export default function manifest(): MetadataRoute.Manifest {
  return {
    name: "Mis Finanzas Diarias",
    short_name: "Mis Finanzas",
    description: "Control personal de ingresos y gastos.",
    start_url: "/",
    display: "standalone",
    background_color: "#fffaf3",
    theme_color: "#c77f73",
    orientation: "portrait-primary",
    lang: "es-PE",
    icons: [
      {
        src: "/favicon.svg",
        sizes: "any",
        type: "image/svg+xml",
        purpose: "any",
      },
    ],
  };
}
