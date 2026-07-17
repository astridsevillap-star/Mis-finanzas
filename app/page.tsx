"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";

type Tipo = "ingreso" | "gasto";
type Movimiento = { id: number; tipo: Tipo; monto: number; concepto: string; categoria: string; fuente: string; detalleFuente?: string; fecha: string };

const categorias = ["Pasajes", "Comida", "Salud", "Pago de deudas", "Suscripciones", "Ropa", "Extras"];
const fuentes = ["Yape", "Plin", "Tarjeta", "Transferencia", "Efectivo", "Otro"];
const colores: Record<string, string> = { Pasajes: "#58c7a2", Comida: "#ff6b5f", Salud: "#9b7bf7", "Pago de deudas": "#3e7bfa", Suscripciones: "#f4b740", Ropa: "#ee7fb2", Extras: "#94a3b8" };
const iniciales: Movimiento[] = [
  { id: 1, tipo: "ingreso", monto: 3200, concepto: "Sueldo", categoria: "Ingresos", fuente: "Transferencia", detalleFuente: "Cuenta bancaria", fecha: "2026-07-15" },
  { id: 2, tipo: "gasto", monto: 86.5, concepto: "Compra semanal", categoria: "Comida", fuente: "Tarjeta", detalleFuente: "Tarjeta principal", fecha: "2026-07-15" },
  { id: 3, tipo: "gasto", monto: 45, concepto: "Recarga de transporte", categoria: "Pasajes", fuente: "Yape", fecha: "2026-07-14" },
  { id: 4, tipo: "gasto", monto: 129.9, concepto: "Suscripción de internet", categoria: "Suscripciones", fuente: "Transferencia", fecha: "2026-07-12" },
];

const soles = (n: number) => new Intl.NumberFormat("es-PE", { style: "currency", currency: "PEN" }).format(n);
const hoy = () => new Date().toISOString().slice(0, 10);

export default function Home() {
  const [movimientos, setMovimientos] = useState<Movimiento[]>(iniciales);
  const [modal, setModal] = useState(false);
  const [tipo, setTipo] = useState<Tipo>("gasto");
  const [monto, setMonto] = useState("");
  const [concepto, setConcepto] = useState("");
  const [categoria, setCategoria] = useState("Comida");
  const [fuente, setFuente] = useState("Yape");
  const [detalleFuente, setDetalleFuente] = useState("");
  const [fecha, setFecha] = useState(hoy());
  const [filtro, setFiltro] = useState("Todos");
  const [filtroFuente, setFiltroFuente] = useState("Todas");
  const [busqueda, setBusqueda] = useState("");
  const [ocultar, setOcultar] = useState(false);
  const [cargado, setCargado] = useState(false);

  useEffect(() => { queueMicrotask(() => { const data = localStorage.getItem("mis-finanzas-movimientos"); if (data) { const mapa: Record<string,string> = { Alimentación: "Comida", Transporte: "Pasajes", Servicios: "Suscripciones", Vivienda: "Extras", Entretenimiento: "Extras", Educación: "Extras", Otros: "Extras" }; setMovimientos((JSON.parse(data) as Movimiento[]).map(m => ({ ...m, categoria: mapa[m.categoria] || m.categoria, fuente: m.fuente || "Otro" }))); } setCargado(true); }); }, []);
  useEffect(() => { if (cargado) localStorage.setItem("mis-finanzas-movimientos", JSON.stringify(movimientos)); }, [movimientos, cargado]);

  const resumen = useMemo(() => {
    const ingresos = movimientos.filter(m => m.tipo === "ingreso").reduce((a, m) => a + m.monto, 0);
    const gastos = movimientos.filter(m => m.tipo === "gasto").reduce((a, m) => a + m.monto, 0);
    const porCategoria = categorias.map(c => ({ nombre: c, total: movimientos.filter(m => m.tipo === "gasto" && m.categoria === c).reduce((a, m) => a + m.monto, 0) })).filter(c => c.total > 0).sort((a, b) => b.total - a.total);
    const porFuente = fuentes.map(f => ({ nombre: f, total: movimientos.filter(m => m.tipo === "gasto" && (m.fuente || "Otro") === f).reduce((a, m) => a + m.monto, 0) })).filter(f => f.total > 0).sort((a,b)=>b.total-a.total);
    return { ingresos, gastos, saldo: ingresos - gastos, porCategoria, porFuente };
  }, [movimientos]);

  const lista = movimientos.filter(m => (filtro === "Todos" || m.tipo === filtro.toLowerCase()) && (filtroFuente === "Todas" || (m.fuente || "Otro") === filtroFuente) && m.concepto.toLowerCase().includes(busqueda.toLowerCase())).sort((a,b) => b.fecha.localeCompare(a.fecha));

  function guardar(e: FormEvent) {
    e.preventDefault(); const valor = Number(monto);
    if (!valor || !concepto.trim()) return;
    setMovimientos(v => [{ id: Date.now(), tipo, monto: valor, concepto: concepto.trim(), categoria: tipo === "ingreso" ? "Ingresos" : categoria, fuente, detalleFuente: detalleFuente.trim(), fecha }, ...v]);
    setMonto(""); setConcepto(""); setDetalleFuente(""); setFecha(hoy()); setModal(false);
  }

  function borrar(id: number) { if (confirm("¿Eliminar este movimiento?")) setMovimientos(v => v.filter(m => m.id !== id)); }

  return <main className="app-shell">
    <aside className="sidebar">
      <div className="brand"><span>MF</span><strong>Mis Finanzas</strong></div>
      <nav>
        <button className="active">⌂ <span>Inicio</span></button><button>⇄ <span>Movimientos</span></button><button>◔ <span>Presupuestos</span></button><button>◎ <span>Metas</span></button><button>▥ <span>Reportes</span></button>
      </nav>
      <div className="tip"><b>✦ Consejo</b><p>Registra incluso los gastos pequeños. Ahí suele esconderse una parte importante de tu dinero.</p></div>
    </aside>

    <section className="content">
      <header><div><p className="eyebrow">DESDE EL 15 DE JUNIO</p><h1>Tu dinero, más claro.</h1><p className="subtitle">Registra tus movimientos y descubre en qué y con qué medio estás pagando.</p></div><button className="avatar">RS</button></header>

      <section className="cards">
        <article className="balance card"><div><span className="label">Balance disponible</span><button className="eye" onClick={() => setOcultar(!ocultar)}>{ocultar ? "◉" : "◌"}</button></div><strong>{ocultar ? "S/ ••••••" : soles(resumen.saldo)}</strong><p>Ingresos menos gastos registrados</p><div className="balance-line" /></article>
        <article className="metric card income"><span className="round">↗</span><div><span className="label">Ingresos</span><strong>{soles(resumen.ingresos)}</strong><p>Acumulado</p></div></article>
        <article className="metric card expense"><span className="round">↘</span><div><span className="label">Gastos</span><strong>{soles(resumen.gastos)}</strong><p>Acumulado</p></div></article>
        <button className="new-button" onClick={() => setModal(true)}><span>＋</span><b>Registrar<br/>movimiento</b><small>Solo toma unos segundos →</small></button>
      </section>

      <section className="lower-grid">
        <article className="card distribution"><div className="section-title"><div><span className="eyebrow">ANÁLISIS</span><h2>¿En qué estás gastando?</h2></div><span className="period">Todos</span></div>
          {resumen.gastos === 0 ? <p className="empty">Aún no hay gastos registrados.</p> : <div className="chart-wrap"><div className="donut" style={{background: `conic-gradient(${resumen.porCategoria.map((c,i) => `${colores[c.nombre]} ${resumen.porCategoria.slice(0,i).reduce((a,x)=>a+x.total,0)/resumen.gastos*100}% ${resumen.porCategoria.slice(0,i+1).reduce((a,x)=>a+x.total,0)/resumen.gastos*100}%`).join(",")})`}}><div><small>Total</small><b>{soles(resumen.gastos)}</b></div></div><div className="legend">{resumen.porCategoria.map(c => <div key={c.nombre}><i style={{background: colores[c.nombre]}}/><span>{c.nombre}</span><b>{Math.round(c.total/resumen.gastos*100)}%</b><small>{soles(c.total)}</small></div>)}</div></div>}
        </article>

        <article className="card movements"><div className="section-title"><div><span className="eyebrow">ACTIVIDAD</span><h2>Movimientos recientes</h2></div><div className="filters"><input aria-label="Buscar" placeholder="Buscar..." value={busqueda} onChange={e=>setBusqueda(e.target.value)}/><select aria-label="Tipo" value={filtro} onChange={e=>setFiltro(e.target.value)}><option>Todos</option><option>Ingreso</option><option>Gasto</option></select><select aria-label="Medio de pago" value={filtroFuente} onChange={e=>setFiltroFuente(e.target.value)}><option>Todas</option>{fuentes.map(f=><option key={f}>{f}</option>)}</select></div></div>
          <div className="source-strip">{resumen.porFuente.slice(0,4).map(f=><div key={f.nombre}><span>{f.nombre}</span><b>{soles(f.total)}</b></div>)}</div>
          <div className="movement-list">{lista.length ? lista.slice(0,8).map(m => <div className="movement" key={m.id}><span className={`movement-icon ${m.tipo}`}>{m.tipo === "ingreso" ? "↓" : "↑"}</span><div className="movement-name"><b>{m.concepto}</b><small>{m.categoria} · {m.fuente || "Otro"}{m.detalleFuente ? ` (${m.detalleFuente})` : ""} · {new Date(m.fecha+"T12:00:00").toLocaleDateString("es-PE", {day:"2-digit",month:"short"})}</small></div><strong className={m.tipo}>{m.tipo === "ingreso" ? "+" : "−"}{soles(m.monto)}</strong><button className="delete" aria-label="Eliminar" onClick={()=>borrar(m.id)}>×</button></div>) : <p className="empty">No encontramos movimientos.</p>}</div>
        </article>
      </section>
    </section>

    {modal && <div className="modal-backdrop" onMouseDown={()=>setModal(false)}><div className="modal" onMouseDown={e=>e.stopPropagation()}><button className="close" onClick={()=>setModal(false)}>×</button><span className="eyebrow">NUEVO REGISTRO</span><h2>Registrar movimiento</h2><p>Puedes cargar movimientos desde el 15 de junio de 2026.</p><form onSubmit={guardar}><div className="type-toggle"><button type="button" className={tipo==="gasto"?"selected gasto":""} onClick={()=>setTipo("gasto")}>Gasto</button><button type="button" className={tipo==="ingreso"?"selected ingreso":""} onClick={()=>setTipo("ingreso")}>Ingreso</button></div><label>Monto (S/)<input autoFocus inputMode="decimal" type="number" min="0.01" step="0.01" placeholder="0.00" value={monto} onChange={e=>setMonto(e.target.value)} required/></label><label>Concepto<input placeholder="Ej. Almuerzo, sueldo, pago de tarjeta" value={concepto} onChange={e=>setConcepto(e.target.value)} required/></label>{tipo==="gasto"&&<label>Categoría<select value={categoria} onChange={e=>setCategoria(e.target.value)}>{categorias.map(c=><option key={c}>{c}</option>)}</select></label>}<label>Medio de pago o fuente<select value={fuente} onChange={e=>setFuente(e.target.value)}>{fuentes.map(f=><option key={f}>{f}</option>)}</select></label><label>{fuente === "Tarjeta" ? "¿Qué tarjeta utilizaste?" : "Detalle de la fuente (opcional)"}<input placeholder={fuente === "Tarjeta" ? "Ej. BCP Visa, Interbank, Ripley" : "Ej. Cuenta BCP, billetera personal"} value={detalleFuente} onChange={e=>setDetalleFuente(e.target.value)}/></label><label>Fecha<input type="date" min="2026-06-15" value={fecha} onChange={e=>setFecha(e.target.value)} required/></label><button className="save" type="submit">Guardar movimiento</button></form></div></div>}
  </main>;
}
