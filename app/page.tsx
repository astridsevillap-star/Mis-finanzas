"use client";

import {
  ChangeEvent,
  FormEvent,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";

type Tipo = "ingreso" | "gasto";
type Movimiento = {
  id: number;
  tipo: Tipo;
  monto: number;
  concepto: string;
  categoria: string;
  fuente: string;
  detalleFuente?: string;
  fecha: string;
};
type MovimientoImportado = Movimiento & { descartado?: boolean };
type FilaExcel = Record<string, unknown>;
type CampoImportacion =
  | "fecha"
  | "concepto"
  | "monto"
  | "cargo"
  | "abono"
  | "tipo"
  | "fuente";
type MapaColumnas = Record<CampoImportacion, string>;

const categorias = [
  "Pasajes",
  "Comida",
  "Salud",
  "Pago de deudas",
  "Suscripciones",
  "Ropa",
  "Extras",
];
const fuentes = [
  "Yape",
  "Plin",
  "Tarjeta",
  "Transferencia",
  "Efectivo",
  "Otro",
];
const colores: Record<string, string> = {
  Pasajes: "#58c7a2",
  Comida: "#ff6b5f",
  Salud: "#9b7bf7",
  "Pago de deudas": "#3e7bfa",
  Suscripciones: "#f4b740",
  Ropa: "#ee7fb2",
  Extras: "#94a3b8",
};
const iniciales: Movimiento[] = [];
const mapaVacio: MapaColumnas = {
  fecha: "",
  concepto: "",
  monto: "",
  cargo: "",
  abono: "",
  tipo: "",
  fuente: "",
};

const soles = (n: number) =>
  new Intl.NumberFormat("es-PE", {
    style: "currency",
    currency: "PEN",
  }).format(n);
const hoy = () => new Date().toISOString().slice(0, 10);
const sinAcentos = (valor: string) =>
  valor
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();

function detectarColumna(columnas: string[], opciones: string[]) {
  const alias = opciones.map(sinAcentos);
  return (
    columnas.find((columna) => {
      const normalizada = sinAcentos(columna);
      if (!normalizada) return false;
      return alias.some(
        (opcion) =>
          normalizada === opcion ||
          normalizada.includes(opcion) ||
          opcion.includes(normalizada),
      );
    }) || ""
  );
}

function detectarMapa(columnas: string[]): MapaColumnas {
  return {
    fecha: detectarColumna(columnas, [
      "fecha",
      "fecha operacion",
      "fecha movimiento",
      "fecha transaccion",
      "date",
    ]),
    concepto: detectarColumna(columnas, [
      "concepto",
      "descripcion",
      "detalle",
      "glosa",
      "operacion",
      "comercio",
      "movimiento",
    ]),
    monto: detectarColumna(columnas, [
      "monto",
      "importe",
      "valor",
      "amount",
      "monto operacion",
    ]),
    cargo: detectarColumna(columnas, [
      "cargo",
      "debito",
      "debe",
      "salida",
      "retiro",
    ]),
    abono: detectarColumna(columnas, [
      "abono",
      "credito",
      "haber",
      "entrada",
      "deposito",
    ]),
    tipo: detectarColumna(columnas, [
      "tipo",
      "tipo movimiento",
      "naturaleza",
    ]),
    fuente: detectarColumna(columnas, [
      "medio",
      "medio de pago",
      "fuente",
      "cuenta",
      "tarjeta",
      "canal",
    ]),
  };
}

function numeroDesdeExcel(valor: unknown): number {
  if (typeof valor === "number") return Number.isFinite(valor) ? valor : 0;
  if (valor === null || valor === undefined) return 0;

  let texto = String(valor).trim();
  const negativoPorParentesis = texto.startsWith("(") && texto.endsWith(")");
  texto = texto.replace(/[^\d,.\-]/g, "");
  if (!texto) return 0;

  const ultimaComa = texto.lastIndexOf(",");
  const ultimoPunto = texto.lastIndexOf(".");
  if (ultimaComa >= 0 && ultimoPunto >= 0) {
    const decimal = ultimaComa > ultimoPunto ? "," : ".";
    const miles = decimal === "," ? /\./g : /,/g;
    texto = texto.replace(miles, "").replace(decimal, ".");
  } else if (ultimaComa >= 0) {
    const decimales = texto.length - ultimaComa - 1;
    texto =
      decimales === 1 || decimales === 2
        ? texto.replace(",", ".")
        : texto.replace(/,/g, "");
  } else if (ultimoPunto >= 0) {
    const decimales = texto.length - ultimoPunto - 1;
    if (decimales === 3 && texto.split(".").length === 2) {
      texto = texto.replace(".", "");
    }
  }

  const numero = Number(texto);
  if (!Number.isFinite(numero)) return 0;
  return negativoPorParentesis ? -Math.abs(numero) : numero;
}

function fechaDesdeExcel(valor: unknown): string {
  if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
    const anio = valor.getFullYear();
    const mes = String(valor.getMonth() + 1).padStart(2, "0");
    const dia = String(valor.getDate()).padStart(2, "0");
    return `${anio}-${mes}-${dia}`;
  }

  if (typeof valor === "number" && valor > 1) {
    const fecha = new Date(Date.UTC(1899, 11, 30) + valor * 86400000);
    return fecha.toISOString().slice(0, 10);
  }

  const texto = String(valor ?? "").trim();
  if (!texto) return "";
  const iso = texto.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) {
    return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;
  }
  const local = texto.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (local) {
    const anio = local[3].length === 2 ? `20${local[3]}` : local[3];
    return `${anio}-${local[2].padStart(2, "0")}-${local[1].padStart(2, "0")}`;
  }
  const fecha = new Date(texto);
  return Number.isNaN(fecha.getTime()) ? "" : fecha.toISOString().slice(0, 10);
}

function leerCsv(textoOriginal: string): unknown[][] {
  const texto = textoOriginal.replace(/^\uFEFF/, "");
  const primeraLinea = texto.split(/\r?\n/, 1)[0] || "";
  const candidatos = [",", ";", "\t"];
  const delimitador = candidatos
    .map((item) => ({
      item,
      total: primeraLinea.split(item).length - 1,
    }))
    .sort((a, b) => b.total - a.total)[0].item;
  const filas: unknown[][] = [];
  let fila: string[] = [];
  let celda = "";
  let entreComillas = false;

  for (let indice = 0; indice < texto.length; indice += 1) {
    const caracter = texto[indice];
    if (entreComillas) {
      if (caracter === '"' && texto[indice + 1] === '"') {
        celda += '"';
        indice += 1;
      } else if (caracter === '"') {
        entreComillas = false;
      } else {
        celda += caracter;
      }
    } else if (caracter === '"') {
      entreComillas = true;
    } else if (caracter === delimitador) {
      fila.push(celda.trim());
      celda = "";
    } else if (caracter === "\n") {
      fila.push(celda.trim());
      if (fila.some((valor) => valor !== "")) filas.push(fila);
      fila = [];
      celda = "";
    } else if (caracter !== "\r") {
      celda += caracter;
    }
  }
  fila.push(celda.trim());
  if (fila.some((valor) => valor !== "")) filas.push(fila);
  return filas;
}

function tipoDesdeTexto(valor: unknown): Tipo | null {
  const texto = sinAcentos(String(valor ?? ""));
  if (/ingreso|abono|credito|haber|entrada/.test(texto)) return "ingreso";
  if (/gasto|cargo|debito|debe|salida|consumo/.test(texto)) return "gasto";
  return null;
}

function sugerirCategoria(concepto: string) {
  const texto = sinAcentos(concepto);
  const reglas: Array<[string, RegExp]> = [
    [
      "Comida",
      /supermerc|plaza vea|vivanda|wong|tottus|metro|mercado|restaur|cafeter|almuerzo|cena|desayuno|rappi|pedidos|panader/,
    ],
    [
      "Pasajes",
      /uber|cabify|indrive|taxi|bus|metropolitano|corredor|transporte|pasaje|grifo|gasolina|combustible|peaje/,
    ],
    [
      "Salud",
      /farmacia|inkafarma|mifarma|botica|clinica|doctor|medico|laboratorio|salud|seguro/,
    ],
    [
      "Pago de deudas",
      /pago.*tarjeta|tarjeta.*pago|prestamo|credito|cuota|deuda|financiamiento/,
    ],
    [
      "Suscripciones",
      /netflix|spotify|internet|movistar|claro|entel|disney|prime|hbo|google|icloud|suscrip|membresia/,
    ],
    [
      "Ropa",
      /ropa|calzado|zapat|falabella|ripley|zara|h&m|oechsle|tienda.*depart/,
    ],
  ];
  return reglas.find(([, patron]) => patron.test(texto))?.[0] || "Extras";
}

function normalizarFuente(valor: unknown) {
  const original = String(valor ?? "").trim();
  const texto = sinAcentos(original);
  if (texto.includes("yape")) return { fuente: "Yape", detalle: "" };
  if (texto.includes("plin")) return { fuente: "Plin", detalle: "" };
  if (/tarjeta|visa|mastercard|amex/.test(texto)) {
    return { fuente: "Tarjeta", detalle: original };
  }
  if (/transfer|cuenta|banco/.test(texto)) {
    return { fuente: "Transferencia", detalle: original };
  }
  if (/efectivo|cash/.test(texto)) return { fuente: "Efectivo", detalle: "" };
  return { fuente: "Otro", detalle: original };
}

function claveMovimiento(movimiento: Movimiento) {
  return [
    movimiento.fecha,
    sinAcentos(movimiento.concepto),
    movimiento.tipo,
    movimiento.monto.toFixed(2),
  ].join("|");
}

export default function Home() {
  const [movimientos, setMovimientos] = useState<Movimiento[]>(iniciales);
  const [modal, setModal] = useState(false);
  const [modalImportar, setModalImportar] = useState(false);
  const [modalRevision, setModalRevision] = useState(false);
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
  const [nombreArchivo, setNombreArchivo] = useState("");
  const [filasExcel, setFilasExcel] = useState<FilaExcel[]>([]);
  const [columnasExcel, setColumnasExcel] = useState<string[]>([]);
  const [mapaColumnas, setMapaColumnas] =
    useState<MapaColumnas>(mapaVacio);
  const [positivosSon, setPositivosSon] = useState<Tipo>("gasto");
  const [errorImportacion, setErrorImportacion] = useState("");
  const [borradores, setBorradores] = useState<MovimientoImportado[]>([]);
  const [indiceRevision, setIndiceRevision] = useState(0);
  const [duplicadosOmitidos, setDuplicadosOmitidos] = useState(0);
  const inputArchivo = useRef<HTMLInputElement>(null);

  useEffect(() => {
    queueMicrotask(() => {
      const data = localStorage.getItem("mis-finanzas-movimientos");
      if (data) {
        const mapa: Record<string, string> = {
          Alimentación: "Comida",
          Transporte: "Pasajes",
          Servicios: "Suscripciones",
          Vivienda: "Extras",
          Entretenimiento: "Extras",
          Educación: "Extras",
          Otros: "Extras",
        };
        setMovimientos(
          (JSON.parse(data) as Movimiento[]).map((movimiento) => ({
            ...movimiento,
            categoria: mapa[movimiento.categoria] || movimiento.categoria,
            fuente: movimiento.fuente || "Otro",
          })),
        );
      }
      setCargado(true);
    });
  }, []);

  useEffect(() => {
    if (cargado) {
      localStorage.setItem(
        "mis-finanzas-movimientos",
        JSON.stringify(movimientos),
      );
    }
  }, [movimientos, cargado]);

  const resumen = useMemo(() => {
    const ingresos = movimientos
      .filter((movimiento) => movimiento.tipo === "ingreso")
      .reduce((total, movimiento) => total + movimiento.monto, 0);
    const gastos = movimientos
      .filter((movimiento) => movimiento.tipo === "gasto")
      .reduce((total, movimiento) => total + movimiento.monto, 0);
    const porCategoria = categorias
      .map((nombre) => ({
        nombre,
        total: movimientos
          .filter(
            (movimiento) =>
              movimiento.tipo === "gasto" &&
              movimiento.categoria === nombre,
          )
          .reduce((total, movimiento) => total + movimiento.monto, 0),
      }))
      .filter((item) => item.total > 0)
      .sort((a, b) => b.total - a.total);
    const porFuente = fuentes
      .map((nombre) => ({
        nombre,
        total: movimientos
          .filter(
            (movimiento) =>
              movimiento.tipo === "gasto" &&
              (movimiento.fuente || "Otro") === nombre,
          )
          .reduce((total, movimiento) => total + movimiento.monto, 0),
      }))
      .filter((item) => item.total > 0)
      .sort((a, b) => b.total - a.total);
    return { ingresos, gastos, saldo: ingresos - gastos, porCategoria, porFuente };
  }, [movimientos]);

  const lista = movimientos
    .filter(
      (movimiento) =>
        (filtro === "Todos" || movimiento.tipo === filtro.toLowerCase()) &&
        (filtroFuente === "Todas" ||
          (movimiento.fuente || "Otro") === filtroFuente) &&
        movimiento.concepto.toLowerCase().includes(busqueda.toLowerCase()),
    )
    .sort((a, b) => b.fecha.localeCompare(a.fecha));

  const actual = borradores[indiceRevision];
  const totalAImportar = borradores.filter(
    (movimiento) => !movimiento.descartado,
  ).length;

  function guardar(e: FormEvent) {
    e.preventDefault();
    const valor = Number(monto);
    if (!valor || !concepto.trim()) return;
    setMovimientos((anteriores) => [
      {
        id: Date.now(),
        tipo,
        monto: valor,
        concepto: concepto.trim(),
        categoria: tipo === "ingreso" ? "Ingresos" : categoria,
        fuente,
        detalleFuente: detalleFuente.trim(),
        fecha,
      },
      ...anteriores,
    ]);
    setMonto("");
    setConcepto("");
    setDetalleFuente("");
    setFecha(hoy());
    setModal(false);
  }

  function borrar(id: number) {
    if (confirm("¿Eliminar este movimiento?")) {
      setMovimientos((anteriores) =>
        anteriores.filter((movimiento) => movimiento.id !== id),
      );
    }
  }

  function reiniciarImportacion() {
    setNombreArchivo("");
    setFilasExcel([]);
    setColumnasExcel([]);
    setMapaColumnas(mapaVacio);
    setErrorImportacion("");
    setBorradores([]);
    setIndiceRevision(0);
    setDuplicadosOmitidos(0);
    if (inputArchivo.current) inputArchivo.current.value = "";
  }

  function cerrarImportacion() {
    setModalImportar(false);
    reiniciarImportacion();
  }

  async function cargarExcel(e: ChangeEvent<HTMLInputElement>) {
    const archivo = e.target.files?.[0];
    if (!archivo) return;
    setErrorImportacion("");

    try {
      let matriz: unknown[][];
      if (archivo.name.toLowerCase().endsWith(".csv")) {
        matriz = leerCsv(await archivo.text());
      } else {
        const { readSheet } = await import("read-excel-file/browser");
        matriz = await readSheet(archivo);
      }
      if (matriz.length < 2) {
        throw new Error("La primera hoja no contiene movimientos.");
      }
      const columnas = matriz[0].map((valor, indice, encabezados) => {
        const base = String(valor ?? "").trim() || `Columna ${indice + 1}`;
        const repeticiones = encabezados
          .slice(0, indice)
          .filter((anterior) => String(anterior ?? "").trim() === base).length;
        return repeticiones ? `${base} (${repeticiones + 1})` : base;
      });
      const filas = matriz
        .slice(1)
        .filter((fila) =>
          fila.some(
            (valor) => valor !== null && valor !== undefined && valor !== "",
          ),
        )
        .map((fila) =>
          Object.fromEntries(
            columnas.map((columna, indice) => [columna, fila[indice] ?? ""]),
          ),
        );
      if (!filas.length) {
        throw new Error("La primera hoja no contiene movimientos.");
      }
      setNombreArchivo(archivo.name);
      setFilasExcel(filas);
      setColumnasExcel(columnas);
      setMapaColumnas(detectarMapa(columnas));
    } catch {
      reiniciarImportacion();
      setErrorImportacion(
        "No pudimos leer el archivo. Verifica que sea un Excel (.xlsx) o CSV válido y que la primera fila contenga los encabezados.",
      );
    }
  }

  function valorFila(fila: FilaExcel, columna: string) {
    return columna ? fila[columna] : "";
  }

  function prepararRevision() {
    setErrorImportacion("");
    if (!mapaColumnas.fecha || !mapaColumnas.concepto) {
      setErrorImportacion(
        "Selecciona las columnas de fecha y descripción del movimiento.",
      );
      return;
    }
    if (
      !mapaColumnas.monto &&
      !mapaColumnas.cargo &&
      !mapaColumnas.abono
    ) {
      setErrorImportacion(
        "Selecciona una columna de monto o las columnas de cargo y abono.",
      );
      return;
    }

    const existentes = new Set(movimientos.map(claveMovimiento));
    const nuevasClaves = new Set<string>();
    let duplicados = 0;
    const baseId = Date.now();
    const importados = filasExcel.flatMap((fila, indice) => {
      const conceptoImportado = String(
        valorFila(fila, mapaColumnas.concepto) ?? "",
      ).trim();
      const fechaImportada = fechaDesdeExcel(
        valorFila(fila, mapaColumnas.fecha),
      );
      if (!conceptoImportado || !fechaImportada) return [];

      const cargo = numeroDesdeExcel(valorFila(fila, mapaColumnas.cargo));
      const abono = numeroDesdeExcel(valorFila(fila, mapaColumnas.abono));
      const montoUnico = numeroDesdeExcel(
        valorFila(fila, mapaColumnas.monto),
      );
      let tipoImportado =
        tipoDesdeTexto(valorFila(fila, mapaColumnas.tipo)) || positivosSon;
      let montoImportado = Math.abs(montoUnico);

      if (mapaColumnas.cargo || mapaColumnas.abono) {
        if (Math.abs(cargo) > 0) {
          tipoImportado = "gasto";
          montoImportado = Math.abs(cargo);
        } else if (Math.abs(abono) > 0) {
          tipoImportado = "ingreso";
          montoImportado = Math.abs(abono);
        }
      } else if (!mapaColumnas.tipo && montoUnico < 0) {
        tipoImportado = "gasto";
      }
      if (!montoImportado) return [];

      const fuenteImportada = normalizarFuente(
        valorFila(fila, mapaColumnas.fuente),
      );
      const movimiento: MovimientoImportado = {
        id: baseId + indice,
        tipo: tipoImportado,
        monto: montoImportado,
        concepto: conceptoImportado,
        categoria:
          tipoImportado === "ingreso"
            ? "Ingresos"
            : sugerirCategoria(conceptoImportado),
        fuente: fuenteImportada.fuente,
        detalleFuente: fuenteImportada.detalle,
        fecha: fechaImportada,
      };
      const clave = claveMovimiento(movimiento);
      if (existentes.has(clave) || nuevasClaves.has(clave)) {
        duplicados += 1;
        return [];
      }
      nuevasClaves.add(clave);
      return [movimiento];
    });

    if (!importados.length) {
      setErrorImportacion(
        duplicados
          ? "Todos los movimientos del archivo ya están registrados."
          : "No encontramos filas válidas. Revisa las columnas seleccionadas y los formatos de fecha y monto.",
      );
      return;
    }
    setBorradores(importados);
    setDuplicadosOmitidos(duplicados);
    setIndiceRevision(0);
    setModalImportar(false);
    setModalRevision(true);
  }

  function actualizarActual(cambios: Partial<MovimientoImportado>) {
    setBorradores((anteriores) =>
      anteriores.map((movimiento, indice) =>
        indice === indiceRevision ? { ...movimiento, ...cambios } : movimiento,
      ),
    );
  }

  function cambiarTipoActual(nuevoTipo: Tipo) {
    actualizarActual({
      tipo: nuevoTipo,
      categoria:
        nuevoTipo === "ingreso"
          ? "Ingresos"
          : actual?.categoria === "Ingresos"
            ? sugerirCategoria(actual.concepto)
            : actual?.categoria,
    });
  }

  function alternarOmitirActual() {
    if (actual.descartado) {
      actualizarActual({ descartado: false });
      return;
    }
    actualizarActual({ descartado: true });
    if (indiceRevision < borradores.length - 1) {
      setIndiceRevision((indice) => indice + 1);
    }
  }

  function guardarImportacion() {
    const confirmados = borradores
      .filter((movimiento) => !movimiento.descartado)
      .map(
        ({
          id,
          tipo: tipoMovimiento,
          monto: montoMovimiento,
          concepto: conceptoMovimiento,
          categoria: categoriaMovimiento,
          fuente: fuenteMovimiento,
          detalleFuente: detalleFuenteMovimiento,
          fecha: fechaMovimiento,
        }): Movimiento => ({
          id,
          tipo: tipoMovimiento,
          monto: montoMovimiento,
          concepto: conceptoMovimiento,
          categoria: categoriaMovimiento,
          fuente: fuenteMovimiento,
          detalleFuente: detalleFuenteMovimiento,
          fecha: fechaMovimiento,
        }),
      );
    if (confirmados.length) {
      setMovimientos((anteriores) => [...confirmados, ...anteriores]);
    }
    setModalRevision(false);
    reiniciarImportacion();
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span>MF</span>
          <strong>Mis Finanzas</strong>
        </div>
        <nav>
          <button className="active">
            ⌂ <span>Inicio</span>
          </button>
          <button>
            ⇄ <span>Movimientos</span>
          </button>
          <button>
            ◔ <span>Presupuestos</span>
          </button>
          <button>
            ◎ <span>Metas</span>
          </button>
          <button>
            ▥ <span>Reportes</span>
          </button>
        </nav>
        <div className="tip">
          <b>✦ Consejo</b>
          <p>
            Registra incluso los gastos pequeños. Ahí suele esconderse una
            parte importante de tu dinero.
          </p>
        </div>
      </aside>

      <section className="content">
        <header>
          <div>
            <p className="eyebrow">DESDE EL 15 DE JUNIO</p>
            <h1>Tu dinero, más claro.</h1>
            <p className="subtitle">
              Registra tus movimientos y descubre en qué y con qué medio estás
              pagando.
            </p>
          </div>
          <button className="avatar">RS</button>
        </header>

        <section className="cards">
          <article className="balance card">
            <div>
              <span className="label">Balance disponible</span>
              <button
                className="eye"
                onClick={() => setOcultar(!ocultar)}
                aria-label={ocultar ? "Mostrar balance" : "Ocultar balance"}
              >
                {ocultar ? "◉" : "◌"}
              </button>
            </div>
            <strong>{ocultar ? "S/ ••••••" : soles(resumen.saldo)}</strong>
            <p>Ingresos menos gastos registrados</p>
            <div className="balance-line" />
          </article>
          <article className="metric card income">
            <span className="round">↗</span>
            <div>
              <span className="label">Ingresos</span>
              <strong>{soles(resumen.ingresos)}</strong>
              <p>Acumulado</p>
            </div>
          </article>
          <article className="metric card expense">
            <span className="round">↘</span>
            <div>
              <span className="label">Gastos</span>
              <strong>{soles(resumen.gastos)}</strong>
              <p>Acumulado</p>
            </div>
          </article>
          <div className="action-stack">
            <button className="new-button compact" onClick={() => setModal(true)}>
              <span>＋</span>
              <b>Registrar movimiento</b>
              <small>Ingreso manual</small>
            </button>
            <button
              className="import-button"
              onClick={() => setModalImportar(true)}
            >
              <span>⇧</span>
              <b>Importar Excel</b>
              <small>.xlsx, .xls o .csv</small>
            </button>
          </div>
        </section>

        <section className="lower-grid">
          <article className="card distribution">
            <div className="section-title">
              <div>
                <span className="eyebrow">ANÁLISIS</span>
                <h2>¿En qué estás gastando?</h2>
              </div>
              <span className="period">Todos</span>
            </div>
            {resumen.gastos === 0 ? (
              <p className="empty">Aún no hay gastos registrados.</p>
            ) : (
              <div className="chart-wrap">
                <div
                  className="donut"
                  style={{
                    background: `conic-gradient(${resumen.porCategoria
                      .map(
                        (item, indice) =>
                          `${colores[item.nombre]} ${
                            (resumen.porCategoria
                              .slice(0, indice)
                              .reduce(
                                (total, categoriaItem) =>
                                  total + categoriaItem.total,
                                0,
                              ) /
                              resumen.gastos) *
                            100
                          }% ${
                            (resumen.porCategoria
                              .slice(0, indice + 1)
                              .reduce(
                                (total, categoriaItem) =>
                                  total + categoriaItem.total,
                                0,
                              ) /
                              resumen.gastos) *
                            100
                          }%`,
                      )
                      .join(",")})`,
                  }}
                >
                  <div>
                    <small>Total</small>
                    <b>{soles(resumen.gastos)}</b>
                  </div>
                </div>
                <div className="legend">
                  {resumen.porCategoria.map((item) => (
                    <div key={item.nombre}>
                      <i style={{ background: colores[item.nombre] }} />
                      <span>{item.nombre}</span>
                      <b>{Math.round((item.total / resumen.gastos) * 100)}%</b>
                      <small>{soles(item.total)}</small>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </article>

          <article className="card movements">
            <div className="section-title">
              <div>
                <span className="eyebrow">ACTIVIDAD</span>
                <h2>Movimientos recientes</h2>
              </div>
              <div className="filters">
                <input
                  aria-label="Buscar"
                  placeholder="Buscar..."
                  value={busqueda}
                  onChange={(e) => setBusqueda(e.target.value)}
                />
                <select
                  aria-label="Tipo"
                  value={filtro}
                  onChange={(e) => setFiltro(e.target.value)}
                >
                  <option>Todos</option>
                  <option>Ingreso</option>
                  <option>Gasto</option>
                </select>
                <select
                  aria-label="Medio de pago"
                  value={filtroFuente}
                  onChange={(e) => setFiltroFuente(e.target.value)}
                >
                  <option>Todas</option>
                  {fuentes.map((item) => (
                    <option key={item}>{item}</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="source-strip">
              {resumen.porFuente.slice(0, 4).map((item) => (
                <div key={item.nombre}>
                  <span>{item.nombre}</span>
                  <b>{soles(item.total)}</b>
                </div>
              ))}
            </div>
            <div className="movement-list">
              {lista.length ? (
                lista.slice(0, 8).map((movimiento) => (
                  <div className="movement" key={movimiento.id}>
                    <span className={`movement-icon ${movimiento.tipo}`}>
                      {movimiento.tipo === "ingreso" ? "↓" : "↑"}
                    </span>
                    <div className="movement-name">
                      <b>{movimiento.concepto}</b>
                      <small>
                        {movimiento.categoria} · {movimiento.fuente || "Otro"}
                        {movimiento.detalleFuente
                          ? ` (${movimiento.detalleFuente})`
                          : ""}{" "}
                        ·{" "}
                        {new Date(
                          `${movimiento.fecha}T12:00:00`,
                        ).toLocaleDateString("es-PE", {
                          day: "2-digit",
                          month: "short",
                        })}
                      </small>
                    </div>
                    <strong className={movimiento.tipo}>
                      {movimiento.tipo === "ingreso" ? "+" : "−"}
                      {soles(movimiento.monto)}
                    </strong>
                    <button
                      className="delete"
                      aria-label="Eliminar"
                      onClick={() => borrar(movimiento.id)}
                    >
                      ×
                    </button>
                  </div>
                ))
              ) : (
                <p className="empty">No encontramos movimientos.</p>
              )}
            </div>
          </article>
        </section>
      </section>

      {modal && (
        <div className="modal-backdrop" onMouseDown={() => setModal(false)}>
          <div className="modal" onMouseDown={(e) => e.stopPropagation()}>
            <button className="close" onClick={() => setModal(false)}>
              ×
            </button>
            <span className="eyebrow">NUEVO REGISTRO</span>
            <h2>Registrar movimiento</h2>
            <p>Puedes cargar movimientos desde el 15 de junio de 2026.</p>
            <form onSubmit={guardar}>
              <div className="type-toggle">
                <button
                  type="button"
                  className={tipo === "gasto" ? "selected gasto" : ""}
                  onClick={() => setTipo("gasto")}
                >
                  Gasto
                </button>
                <button
                  type="button"
                  className={tipo === "ingreso" ? "selected ingreso" : ""}
                  onClick={() => setTipo("ingreso")}
                >
                  Ingreso
                </button>
              </div>
              <label>
                Monto (S/)
                <input
                  autoFocus
                  inputMode="decimal"
                  type="number"
                  min="0.01"
                  step="0.01"
                  placeholder="0.00"
                  value={monto}
                  onChange={(e) => setMonto(e.target.value)}
                  required
                />
              </label>
              <label>
                Concepto
                <input
                  placeholder="Ej. Almuerzo, sueldo, pago de tarjeta"
                  value={concepto}
                  onChange={(e) => setConcepto(e.target.value)}
                  required
                />
              </label>
              {tipo === "gasto" && (
                <label>
                  Categoría
                  <select
                    value={categoria}
                    onChange={(e) => setCategoria(e.target.value)}
                  >
                    {categorias.map((item) => (
                      <option key={item}>{item}</option>
                    ))}
                  </select>
                </label>
              )}
              <label>
                Medio de pago o fuente
                <select
                  value={fuente}
                  onChange={(e) => setFuente(e.target.value)}
                >
                  {fuentes.map((item) => (
                    <option key={item}>{item}</option>
                  ))}
                </select>
              </label>
              <label>
                {fuente === "Tarjeta"
                  ? "¿Qué tarjeta utilizaste?"
                  : "Detalle de la fuente (opcional)"}
                <input
                  placeholder={
                    fuente === "Tarjeta"
                      ? "Ej. BCP Visa, Interbank, Ripley"
                      : "Ej. Cuenta BCP, billetera personal"
                  }
                  value={detalleFuente}
                  onChange={(e) => setDetalleFuente(e.target.value)}
                />
              </label>
              <label>
                Fecha
                <input
                  type="date"
                  min="2026-06-15"
                  value={fecha}
                  onChange={(e) => setFecha(e.target.value)}
                  required
                />
              </label>
              <button className="save" type="submit">
                Guardar movimiento
              </button>
            </form>
          </div>
        </div>
      )}

      {modalImportar && (
        <div className="modal-backdrop" onMouseDown={cerrarImportacion}>
          <div
            className="modal import-modal"
            onMouseDown={(e) => e.stopPropagation()}
          >
            <button className="close" onClick={cerrarImportacion}>
              ×
            </button>
            <span className="eyebrow">CARGA MASIVA</span>
            <h2>Importar movimientos</h2>
            <p>
              El archivo se procesa en este dispositivo. Selecciona la primera
              hoja y confirma qué representa cada columna.
            </p>

            <label className={`file-drop ${nombreArchivo ? "loaded" : ""}`}>
              <input
                ref={inputArchivo}
                type="file"
                accept=".xlsx,.csv"
                onChange={cargarExcel}
              />
              <span>{nombreArchivo ? "✓" : "⇧"}</span>
              <b>{nombreArchivo || "Seleccionar Excel o CSV"}</b>
              <small>
                {filasExcel.length
                  ? `${filasExcel.length} filas encontradas`
                  : "Formatos permitidos: .xlsx y .csv"}
              </small>
            </label>

            {filasExcel.length > 0 && (
              <>
                <div className="mapping-grid">
                  <label>
                    Fecha *
                    <select
                      value={mapaColumnas.fecha}
                      onChange={(e) =>
                        setMapaColumnas({
                          ...mapaColumnas,
                          fecha: e.target.value,
                        })
                      }
                    >
                      <option value="">Seleccionar columna</option>
                      {columnasExcel.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Descripción *
                    <select
                      value={mapaColumnas.concepto}
                      onChange={(e) =>
                        setMapaColumnas({
                          ...mapaColumnas,
                          concepto: e.target.value,
                        })
                      }
                    >
                      <option value="">Seleccionar columna</option>
                      {columnasExcel.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Monto único
                    <select
                      value={mapaColumnas.monto}
                      onChange={(e) =>
                        setMapaColumnas({
                          ...mapaColumnas,
                          monto: e.target.value,
                        })
                      }
                    >
                      <option value="">No aplica</option>
                      {columnasExcel.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Cargo / débito
                    <select
                      value={mapaColumnas.cargo}
                      onChange={(e) =>
                        setMapaColumnas({
                          ...mapaColumnas,
                          cargo: e.target.value,
                        })
                      }
                    >
                      <option value="">No aplica</option>
                      {columnasExcel.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Abono / crédito
                    <select
                      value={mapaColumnas.abono}
                      onChange={(e) =>
                        setMapaColumnas({
                          ...mapaColumnas,
                          abono: e.target.value,
                        })
                      }
                    >
                      <option value="">No aplica</option>
                      {columnasExcel.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Tipo
                    <select
                      value={mapaColumnas.tipo}
                      onChange={(e) =>
                        setMapaColumnas({
                          ...mapaColumnas,
                          tipo: e.target.value,
                        })
                      }
                    >
                      <option value="">No aplica</option>
                      {columnasExcel.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Medio o cuenta
                    <select
                      value={mapaColumnas.fuente}
                      onChange={(e) =>
                        setMapaColumnas({
                          ...mapaColumnas,
                          fuente: e.target.value,
                        })
                      }
                    >
                      <option value="">No aplica</option>
                      {columnasExcel.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  {!mapaColumnas.tipo &&
                    !mapaColumnas.cargo &&
                    !mapaColumnas.abono && (
                      <label>
                        Los montos positivos son
                        <select
                          value={positivosSon}
                          onChange={(e) =>
                            setPositivosSon(e.target.value as Tipo)
                          }
                        >
                          <option value="gasto">Gastos</option>
                          <option value="ingreso">Ingresos</option>
                        </select>
                      </label>
                    )}
                </div>

                <div className="excel-preview">
                  <b>Vista previa</b>
                  {filasExcel.slice(0, 3).map((fila, indice) => (
                    <div key={indice}>
                      <span>
                        {String(
                          valorFila(fila, mapaColumnas.fecha) || "Sin fecha",
                        )}
                      </span>
                      <strong>
                        {String(
                          valorFila(fila, mapaColumnas.concepto) ||
                            "Sin descripción",
                        )}
                      </strong>
                      <span>
                        {String(
                          valorFila(
                            fila,
                            mapaColumnas.monto ||
                              mapaColumnas.cargo ||
                              mapaColumnas.abono,
                          ) || "Sin monto",
                        )}
                      </span>
                    </div>
                  ))}
                </div>
              </>
            )}

            {errorImportacion && (
              <p className="import-error">{errorImportacion}</p>
            )}
            <button
              className="save"
              type="button"
              disabled={!filasExcel.length}
              onClick={prepararRevision}
            >
              Revisar movimientos uno a uno
            </button>
          </div>
        </div>
      )}

      {modalRevision && actual && (
        <div className="modal-backdrop">
          <div className="modal review-modal">
            <span className="eyebrow">REVISIÓN DE CATEGORÍAS</span>
            <div className="review-heading">
              <div>
                <h2>
                  Movimiento {indiceRevision + 1} de {borradores.length}
                </h2>
                <p>
                  Confirma el tipo y la categoría antes de guardar el archivo.
                </p>
              </div>
              <strong>{Math.round(((indiceRevision + 1) / borradores.length) * 100)}%</strong>
            </div>
            <div className="review-progress">
              <i
                style={{
                  width: `${((indiceRevision + 1) / borradores.length) * 100}%`,
                }}
              />
            </div>

            {duplicadosOmitidos > 0 && indiceRevision === 0 && (
              <p className="duplicate-note">
                {duplicadosOmitidos}{" "}
                {duplicadosOmitidos === 1
                  ? "duplicado fue omitido"
                  : "duplicados fueron omitidos"}{" "}
                automáticamente.
              </p>
            )}

            <article
              className={`review-card ${actual.descartado ? "discarded" : ""}`}
            >
              <div className="review-date">
                <span>
                  {new Date(`${actual.fecha}T12:00:00`).toLocaleDateString(
                    "es-PE",
                    { day: "2-digit", month: "long", year: "numeric" },
                  )}
                </span>
                <strong>{soles(actual.monto)}</strong>
              </div>
              <h3>{actual.concepto}</h3>
              {actual.detalleFuente && <p>{actual.detalleFuente}</p>}
              <div className="type-toggle">
                <button
                  type="button"
                  className={actual.tipo === "gasto" ? "selected gasto" : ""}
                  onClick={() => cambiarTipoActual("gasto")}
                >
                  Gasto
                </button>
                <button
                  type="button"
                  className={
                    actual.tipo === "ingreso" ? "selected ingreso" : ""
                  }
                  onClick={() => cambiarTipoActual("ingreso")}
                >
                  Ingreso
                </button>
              </div>
              {actual.tipo === "gasto" ? (
                <label>
                  ¿A qué categoría corresponde?
                  <select
                    value={actual.categoria}
                    onChange={(e) =>
                      actualizarActual({ categoria: e.target.value })
                    }
                  >
                    {categorias.map((item) => (
                      <option key={item}>{item}</option>
                    ))}
                  </select>
                </label>
              ) : (
                <div className="income-category">
                  <span>Categoría</span>
                  <b>Ingresos</b>
                </div>
              )}
              <label>
                Medio de pago o fuente
                <select
                  value={actual.fuente}
                  onChange={(e) =>
                    actualizarActual({ fuente: e.target.value })
                  }
                >
                  {fuentes.map((item) => (
                    <option key={item}>{item}</option>
                  ))}
                </select>
              </label>
              {actual.descartado && (
                <p className="discarded-note">
                  Este movimiento no se guardará.
                </p>
              )}
            </article>

            <div className="review-actions">
              <button
                type="button"
                className="secondary"
                disabled={indiceRevision === 0}
                onClick={() => setIndiceRevision((indice) => indice - 1)}
              >
                ← Anterior
              </button>
              <button
                type="button"
                className="skip"
                onClick={alternarOmitirActual}
              >
                {actual.descartado ? "Incluir" : "Omitir"}
              </button>
              {indiceRevision < borradores.length - 1 ? (
                <button
                  type="button"
                  className="primary"
                  onClick={() => setIndiceRevision((indice) => indice + 1)}
                >
                  Siguiente →
                </button>
              ) : (
                <button
                  type="button"
                  className="primary"
                  disabled={!totalAImportar}
                  onClick={guardarImportacion}
                >
                  Guardar {totalAImportar} movimientos
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </main>
  );
}
