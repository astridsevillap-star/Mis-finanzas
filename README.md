# Mis Finanzas Diarias

Aplicación web personal para registrar ingresos y gastos, clasificarlos por categoría y conocer el medio de pago utilizado.

## Funciones incluidas

- Registro de ingresos y gastos desde el 15 de junio de 2026.
- Categorías: Pasajes, Comida, Salud, Pago de deudas, Suscripciones, Ropa y Extras.
- Fuentes: Yape, Plin, Tarjeta, Transferencia, Efectivo y Otro.
- Detalle de la tarjeta o cuenta utilizada.
- Balance automático, distribución por categoría y resumen por medio de pago.
- Búsqueda, filtros y eliminación de movimientos.
- Diseño adaptable a computadora y celular.

## Ejecutar en tu computadora

Necesitas Node.js 22 o superior.

```bash
npm install
npm run dev
```

Luego abre `http://localhost:3000`.

## Subir a GitHub

1. Crea un repositorio nuevo y vacío en GitHub.
2. Descomprime este proyecto.
3. Abre una terminal dentro de la carpeta.
4. Ejecuta:

```bash
git init
git add .
git commit -m "Primera versión de Mis Finanzas"
git branch -M main
git remote add origin URL_DE_TU_REPOSITORIO
git push -u origin main
```

## Publicar en Vercel

1. Ingresa a Vercel y selecciona **Add New > Project**.
2. Importa el repositorio de GitHub.
3. Vercel detectará automáticamente Next.js.
4. Presiona **Deploy**. No necesitas configurar variables de entorno para esta versión.

## Importante sobre los datos

Esta primera versión guarda los movimientos en `localStorage`: permanecen en el navegador y dispositivo donde fueron registrados. No subas estados de cuenta ni información bancaria al repositorio.

Para sincronizar datos entre dispositivos, el siguiente paso recomendado es añadir una base de datos como Supabase, Neon o Vercel Postgres y un sistema de inicio de sesión.
