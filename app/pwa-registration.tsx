"use client";

import { useEffect } from "react";

export default function PwaRegistration() {
  useEffect(() => {
    if ("serviceWorker" in navigator && process.env.NODE_ENV === "production") {
      navigator.serviceWorker.register("/sw.js").catch(() => {
        // La aplicación sigue funcionando en línea si el navegador no admite PWA.
      });
    }
  }, []);

  return null;
}
