"use client";

import {
  ChangeEvent,
  FormEvent,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import type { Session } from "@supabase/supabase-js";
import { supabase } from "@/lib/supabase";

type Tipo = "ingreso" | "gasto" | "transferencia";
type Empresa = string;
type Movimiento = {
  id: number;
  tipo: Tipo;
  monto: number;
  concepto: string;
  categoria: string;
  fuente: string;
  detalleFuente?: string;
  fecha: string;
  empresa: Empresa;
  subcategoria?: string;
  empresaOrigen?: Empresa;
  empresaDestino?: Empresa;
  cuentaOrigen?: string;
  cuentaDestino?: string;
  claseTransferencia?: string;
  excluirPresupuesto?: boolean;
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
type ArchivoImportado = {
  id: string;
  nombre: string;
  filas: FilaExcel[];
  columnas: string[];
  mapa: MapaColumnas;
  positivosSon: Tipo;
};

const categoriasDefault = [
  "Pasajes",
  "Comida",
  "Salud",
  "Pago de deudas",
  "Compras del negocio",
  "Suscripciones",
  "Ropa",
  "Extras",
];
const empresasDefault = ["Personal", "SS", "Clever"];
const subcategoriasExtrasDefault = ["Hogar", "Educación", "Regalos", "Ocio", "Trámites", "Imprevistos", "Otros"];
const categoriasIngresoDefault = ["Sueldo", "Ventas o servicios", "Rendimientos de inversiones", "Aporte para gastos compartidos", "Aporte para ahorro", "Devolución de préstamo", "Otros ingresos", "Saldo inicial"];
const fuentesDefault = [
  "Yape",
  "Plin",
  "Tarjeta",
  "Transferencia",
  "SIP",
  "Efectivo",
  "Otro",
];
const cuentasPorEmpresaDefault: Record<string, string[]> = {
  Personal: ["BCP", "IBK", "Yape", "Plin", "Efectivo", "SIP", "Préstamos a Jair", "Tarjeta", "Otro"],
  SS: ["Cuenta SS", "Efectivo SS", "Otro"],
  Clever: ["Cuenta Clever", "Efectivo Clever", "Otro"],
};
const paletaCategorias = ["#58c7a2", "#ff6b5f", "#9b7bf7", "#3e7bfa", "#f4b740", "#ee7fb2", "#94a3b8", "#7b91b8", "#e6a23c", "#5cb8b2", "#c17ed1", "#7fa8d9"];
function colorCategoria(nombre: string, indice: number) {
  return colores[nombre] || paletaCategorias[indice % paletaCategorias.length];
}
const clasesTransferencia = ["Entre cuentas propias", "Aporte de capital", "Devolución de aporte", "Devolución de préstamo", "Distribución de utilidades", "Ahorro o inversión"];
const colores: Record<string, string> = {
  Pasajes: "#58c7a2",
  Comida: "#ff6b5f",
  Salud: "#9b7bf7",
  "Pago de deudas": "#3e7bfa",
  Suscripciones: "#f4b740",
  Ropa: "#ee7fb2",
  Extras: "#94a3b8",
  "Compras del negocio": "#7b91b8",
};
const iniciales: Movimiento[] = [];
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

function detalleAutomatico(fuente: string) {
  const automaticos: Record<string, string> = {
    Yape: "BCP",
    Transferencia: "BCP",
    Plin: "IBK",
    Tarjeta: "Tarjeta",
    SIP: "SIP",
  };
  return automaticos[fuente] || "";
}

function cuentaMovimiento(movimiento: Movimiento) {
  const detalle = movimiento.detalleFuente?.trim();
  return detalle && ["Yape", "Plin", "Transferencia"].includes(movimiento.fuente)
    ? detalle
    : movimiento.fuente || "Otro";
}

function esPrestamoJair(movimiento: Movimiento) {
  return movimiento.tipo === "gasto" && sinAcentos(movimiento.concepto).includes("jair");
}

function esMovimientoJair(movimiento: Movimiento) {
  return esPrestamoJair(movimiento) ||
    (movimiento.tipo === "transferencia" &&
      (movimiento.claseTransferencia === "Devolución de préstamo" ||
        movimiento.cuentaOrigen === "Préstamos a Jair" ||
        sinAcentos(movimiento.concepto).includes("jair")));
}

function normalizarDevolucionJair(movimiento: Movimiento): Movimiento {
  const esDevolucionRegistrada = movimiento.tipo !== "gasto" && movimiento.fecha === "2026-07-03" && movimiento.monto === 500 && sinAcentos(movimiento.concepto).includes("jair");
  if (!esDevolucionRegistrada) return movimiento;
  const cuentaRecibida = movimiento.detalleFuente || cuentaMovimiento(movimiento) || "BCP";
  return {
    ...movimiento,
    tipo: "transferencia",
    categoria: "Devolución de préstamo",
    fuente: "Préstamos a Jair",
    detalleFuente: cuentaRecibida,
    empresa: "Personal",
    empresaOrigen: "Personal",
    empresaDestino: "Personal",
    cuentaOrigen: "Préstamos a Jair",
    cuentaDestino: cuentaRecibida,
    claseTransferencia: "Devolución de préstamo",
  };
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
  const [empresa, setEmpresa] = useState<Empresa>("Personal");
  const [subcategoria, setSubcategoria] = useState("Otros");
  const [categoriaIngreso, setCategoriaIngreso] = useState("Sueldo");
  const [fuente, setFuente] = useState("Yape");
  const [detalleFuente, setDetalleFuente] = useState("");
  const [cuentaOrigen, setCuentaOrigen] = useState("BCP");
  const [cuentaDestino, setCuentaDestino] = useState("IBK");
  const [empresaOrigen, setEmpresaOrigen] = useState<Empresa>("Personal");
  const [empresaDestino, setEmpresaDestino] = useState<Empresa>("Personal");
  const [claseTransferencia, setClaseTransferencia] = useState("Entre cuentas propias");
  const [fecha, setFecha] = useState(hoy());
  const [filtro, setFiltro] = useState("Todos");
  const [filtroFuente, setFiltroFuente] = useState("Todas");
  const [filtroMes, setFiltroMes] = useState("Todos");
  const [filtroEmpresa, setFiltroEmpresa] = useState("Todas");
  const [soloJair, setSoloJair] = useState(false);
  const [vista, setVista] = useState<"inicio" | "movimientos" | "metas" | "presupuestos" | "reportes" | "ajustes">("inicio");
  const [presupuestos, setPresupuestos] = useState<Record<string, number>>({});
  const [presupuestosCargados, setPresupuestosCargados] = useState(false);
  const [excluirPresupuesto, setExcluirPresupuesto] = useState(false);
  const [busqueda, setBusqueda] = useState("");
  const [ocultar, setOcultar] = useState(false);
  const [extrasAbierto, setExtrasAbierto] = useState<string | null>(null);
  const [cargado, setCargado] = useState(false);
  const [archivosExcel, setArchivosExcel] = useState<ArchivoImportado[]>([]);
  const [indiceArchivo, setIndiceArchivo] = useState(0);
  const [errorImportacion, setErrorImportacion] = useState("");
  const [borradores, setBorradores] = useState<MovimientoImportado[]>([]);
  const [indiceRevision, setIndiceRevision] = useState(0);
  const [duplicadosOmitidos, setDuplicadosOmitidos] = useState(0);
  const [session, setSession] = useState<Session | null>(null);
  const [authRevisada, setAuthRevisada] = useState(false);
  const [modalCuenta, setModalCuenta] = useState(false);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [mensajeCuenta, setMensajeCuenta] = useState("");
  const [procesandoCuenta, setProcesandoCuenta] = useState(false);
  const [nubeLista, setNubeLista] = useState(false);
  const [estadoNube, setEstadoNube] = useState("Guardado en este dispositivo");
  const [configuracionLista, setConfiguracionLista] = useState(false);
  const [categoriasGastoState, setCategoriasGastoState] = useState<string[]>([]);
  const [categoriasIngresoState, setCategoriasIngresoState] = useState<string[]>([]);
  const [subcategoriasExtrasState, setSubcategoriasExtrasState] = useState<string[]>([]);
  const [fuentesState, setFuentesState] = useState<string[]>([]);
  const [empresasState, setEmpresasState] = useState<string[]>([]);
  const [cuentasPorEmpresaState, setCuentasPorEmpresaState] = useState<Record<string, string[]>>({});
  const [nuevaCategoria, setNuevaCategoria] = useState("");
  const [nuevaCategoriaIngreso, setNuevaCategoriaIngreso] = useState("");
  const [nuevaSubcategoria, setNuevaSubcategoria] = useState("");
  const [nuevaFuente, setNuevaFuente] = useState("");
  const [nuevaEmpresa, setNuevaEmpresa] = useState("");
  const [empresaAjustes, setEmpresaAjustes] = useState("");
  const [nuevaCuenta, setNuevaCuenta] = useState("");
  const [movimientoEditando, setMovimientoEditando] = useState<number | null>(null);
  const inputArchivo = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!authRevisada || !session) return;
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
          (JSON.parse(data) as Movimiento[]).map((movimiento) => normalizarDevolucionJair({
            ...movimiento,
            categoria: mapa[movimiento.categoria] || movimiento.categoria,
            fuente: movimiento.fuente || "Otro",
            empresa: movimiento.empresa || "Personal",
            subcategoria: movimiento.subcategoria || "",
            empresaOrigen: movimiento.empresaOrigen || (movimiento.tipo !== "ingreso" ? movimiento.empresa || "Personal" : undefined),
            empresaDestino: movimiento.empresaDestino || (movimiento.tipo !== "gasto" ? movimiento.empresa || "Personal" : undefined),
            cuentaOrigen: movimiento.cuentaOrigen || (movimiento.tipo !== "ingreso" ? cuentaMovimiento(movimiento) : undefined),
            cuentaDestino: movimiento.cuentaDestino || (movimiento.tipo === "transferencia" ? movimiento.detalleFuente : movimiento.tipo === "ingreso" ? cuentaMovimiento(movimiento) : undefined),
            claseTransferencia: movimiento.claseTransferencia || (movimiento.tipo === "transferencia" ? "Entre cuentas propias" : undefined),
            excluirPresupuesto: movimiento.excluirPresupuesto || false,
          })),
        );
      }
      setCargado(true);
    });
  }, [authRevisada, session]);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session);
      setAuthRevisada(true);
    });
    const { data } = supabase.auth.onAuthStateChange((_event, nuevaSession) => {
      setSession(nuevaSession);
      setAuthRevisada(true);
      if (!nuevaSession) {
        setMovimientos([]);
        setPresupuestos({});
        setCargado(false);
        setPresupuestosCargados(false);
        setConfiguracionLista(false);
        setNubeLista(false);
        setEstadoNube("Guardado en este dispositivo");
      }
    });
    return () => data.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session || !cargado || nubeLista) return;
    let activo = true;
    async function cargarNube() {
      setEstadoNube("Sincronizando con la nube…");
      const { data, error } = await supabase
        .from("movimientos")
        .select("client_id,tipo,monto,concepto,categoria,fuente,detalle_fuente,fecha,empresa,subcategoria,empresa_origen,empresa_destino,cuenta_origen,cuenta_destino,clase_transferencia,excluir_presupuesto")
        .order("fecha", { ascending: false });
      if (!activo) return;
      if (error) {
        setEstadoNube("No se pudo conectar con la nube");
        return;
      }
      const remotos: Movimiento[] = (data || []).map((item) => normalizarDevolucionJair({
        id: Number(item.client_id),
        tipo: item.tipo as Tipo,
        monto: Number(item.monto),
        concepto: item.concepto,
        categoria: item.categoria,
        fuente: item.fuente,
        detalleFuente: item.detalle_fuente || "",
        fecha: item.fecha,
        empresa: item.empresa || "Personal",
        subcategoria: item.subcategoria || "",
        empresaOrigen: item.empresa_origen || undefined,
        empresaDestino: item.empresa_destino || undefined,
        cuentaOrigen: item.cuenta_origen || undefined,
        cuentaDestino: item.cuenta_destino || undefined,
        claseTransferencia: item.clase_transferencia || undefined,
        excluirPresupuesto: item.excluir_presupuesto || false,
      }));
      setMovimientos((locales) => {
        const combinados = new Map<number, Movimiento>();
        locales.forEach((item) => combinados.set(item.id, item));
        remotos.forEach((item) => combinados.set(item.id, item));
        return Array.from(combinados.values());
      });
      setNubeLista(true);
      setEstadoNube("Guardado en la nube");
    }
    cargarNube();
    return () => { activo = false; };
  }, [session, cargado, nubeLista]);

  useEffect(() => {
    if (!session) return;
    let activo = true;
    async function cargarConfiguracion() {
      const [categoriasRes, subcategoriasRes, fuentesRes, empresasRes, cuentasRes] = await Promise.all([
        supabase.from("categorias").select("nombre,tipo").order("orden"),
        supabase.from("subcategorias_extras").select("nombre").order("orden"),
        supabase.from("fuentes").select("nombre").order("orden"),
        supabase.from("empresas").select("nombre").order("orden"),
        supabase.from("cuentas").select("empresa,nombre").order("orden"),
      ]);
      if (!activo) return;
      const categoriasFilas = categoriasRes.data || [];
      setCategoriasGastoState(categoriasFilas.filter((item) => item.tipo === "gasto").map((item) => item.nombre));
      setCategoriasIngresoState(categoriasFilas.filter((item) => item.tipo === "ingreso").map((item) => item.nombre));
      setSubcategoriasExtrasState((subcategoriasRes.data || []).map((item) => item.nombre));
      setFuentesState((fuentesRes.data || []).map((item) => item.nombre));
      setEmpresasState((empresasRes.data || []).map((item) => item.nombre));
      const cuentasAgrupadas: Record<string, string[]> = {};
      (cuentasRes.data || []).forEach((item) => {
        if (!cuentasAgrupadas[item.empresa]) cuentasAgrupadas[item.empresa] = [];
        cuentasAgrupadas[item.empresa].push(item.nombre);
      });
      setCuentasPorEmpresaState(cuentasAgrupadas);
      setConfiguracionLista(true);
    }
    cargarConfiguracion();
    return () => { activo = false; };
  }, [session]);

  async function agregarElemento(
    tabla: "categorias" | "subcategorias_extras" | "fuentes" | "empresas" | "cuentas",
    valores: Record<string, unknown>,
    actualizarLocal: () => void,
  ) {
    if (!session) return;
    const { error } = await supabase.from(tabla).insert({ ...valores, user_id: session.user.id });
    if (!error) actualizarLocal();
  }

  async function eliminarElemento(
    tabla: "categorias" | "subcategorias_extras" | "fuentes" | "empresas" | "cuentas",
    condicion: Record<string, unknown>,
    actualizarLocal: () => void,
  ) {
    if (!session) return;
    let consulta = supabase.from(tabla).delete().eq("user_id", session.user.id);
    Object.entries(condicion).forEach(([columna, valor]) => {
      consulta = consulta.eq(columna, valor as string);
    });
    const { error } = await consulta;
    if (!error) actualizarLocal();
  }

  useEffect(() => {
    if (empresasState.length && !empresasState.includes(empresaAjustes)) {
      queueMicrotask(() => setEmpresaAjustes(empresasState[0]));
    }
  }, [empresasState, empresaAjustes]);

  function agregarCategoriaGasto() {
    const nombre = nuevaCategoria.trim();
    if (!nombre || categoriasGastoState.includes(nombre)) return;
    agregarElemento("categorias", { nombre, tipo: "gasto", orden: categoriasGastoState.length }, () => {
      setCategoriasGastoState((anteriores) => [...anteriores, nombre]);
      setNuevaCategoria("");
    });
  }
  function eliminarCategoriaGasto(nombre: string) {
    eliminarElemento("categorias", { nombre, tipo: "gasto" }, () => {
      setCategoriasGastoState((anteriores) => anteriores.filter((item) => item !== nombre));
    });
  }
  function agregarCategoriaIngresoNueva() {
    const nombre = nuevaCategoriaIngreso.trim();
    if (!nombre || categoriasIngresoState.includes(nombre)) return;
    agregarElemento("categorias", { nombre, tipo: "ingreso", orden: categoriasIngresoState.length }, () => {
      setCategoriasIngresoState((anteriores) => [...anteriores, nombre]);
      setNuevaCategoriaIngreso("");
    });
  }
  function eliminarCategoriaIngresoNueva(nombre: string) {
    eliminarElemento("categorias", { nombre, tipo: "ingreso" }, () => {
      setCategoriasIngresoState((anteriores) => anteriores.filter((item) => item !== nombre));
    });
  }
  function agregarSubcategoriaNueva() {
    const nombre = nuevaSubcategoria.trim();
    if (!nombre || subcategoriasExtrasState.includes(nombre)) return;
    agregarElemento("subcategorias_extras", { nombre, orden: subcategoriasExtrasState.length }, () => {
      setSubcategoriasExtrasState((anteriores) => [...anteriores, nombre]);
      setNuevaSubcategoria("");
    });
  }
  function eliminarSubcategoriaNueva(nombre: string) {
    eliminarElemento("subcategorias_extras", { nombre }, () => {
      setSubcategoriasExtrasState((anteriores) => anteriores.filter((item) => item !== nombre));
    });
  }
  function agregarFuenteNueva() {
    const nombre = nuevaFuente.trim();
    if (!nombre || fuentesState.includes(nombre)) return;
    agregarElemento("fuentes", { nombre, orden: fuentesState.length }, () => {
      setFuentesState((anteriores) => [...anteriores, nombre]);
      setNuevaFuente("");
    });
  }
  function eliminarFuenteNueva(nombre: string) {
    eliminarElemento("fuentes", { nombre }, () => {
      setFuentesState((anteriores) => anteriores.filter((item) => item !== nombre));
    });
  }
  function agregarEmpresaNueva() {
    const nombre = nuevaEmpresa.trim();
    if (!nombre || empresasState.includes(nombre)) return;
    agregarElemento("empresas", { nombre, orden: empresasState.length }, () => {
      setEmpresasState((anteriores) => [...anteriores, nombre]);
      setNuevaEmpresa("");
      setEmpresaAjustes(nombre);
    });
  }
  function eliminarEmpresaNueva(nombre: string) {
    eliminarElemento("empresas", { nombre }, () => {
      setEmpresasState((anteriores) => anteriores.filter((item) => item !== nombre));
      setCuentasPorEmpresaState((anteriores) => {
        const copia = { ...anteriores };
        delete copia[nombre];
        return copia;
      });
    });
  }
  function agregarCuentaNueva() {
    const nombre = nuevaCuenta.trim();
    if (!nombre || !empresaAjustes) return;
    const actuales = cuentasPorEmpresaState[empresaAjustes] || [];
    if (actuales.includes(nombre)) return;
    agregarElemento("cuentas", { nombre, empresa: empresaAjustes, orden: actuales.length }, () => {
      setCuentasPorEmpresaState((anteriores) => ({ ...anteriores, [empresaAjustes]: [...(anteriores[empresaAjustes] || []), nombre] }));
      setNuevaCuenta("");
    });
  }
  function eliminarCuentaNueva(nombre: string) {
    eliminarElemento("cuentas", { nombre, empresa: empresaAjustes }, () => {
      setCuentasPorEmpresaState((anteriores) => ({ ...anteriores, [empresaAjustes]: (anteriores[empresaAjustes] || []).filter((item) => item !== nombre) }));
    });
  }

  useEffect(() => {
    if (!session || !nubeLista) return;
    const temporizador = window.setTimeout(async () => {
      setEstadoNube("Guardando en la nube…");
      const filas = movimientos.map((item) => ({
        user_id: session.user.id,
        client_id: item.id,
        tipo: item.tipo,
        monto: item.monto,
        concepto: item.concepto,
        categoria: item.categoria,
        fuente: item.fuente || "Otro",
        detalle_fuente: item.detalleFuente || null,
        fecha: item.fecha,
        empresa: item.empresa || "Personal",
        subcategoria: item.subcategoria || null,
        empresa_origen: item.empresaOrigen || (item.tipo !== "ingreso" ? item.empresa : null),
        empresa_destino: item.empresaDestino || (item.tipo !== "gasto" ? item.empresa : null),
        cuenta_origen: item.cuentaOrigen || (item.tipo !== "ingreso" ? cuentaMovimiento(item) : null),
        cuenta_destino: item.cuentaDestino || (item.tipo === "transferencia" ? item.detalleFuente || null : item.tipo === "ingreso" ? cuentaMovimiento(item) : null),
        clase_transferencia: item.claseTransferencia || null,
        excluir_presupuesto: item.excluirPresupuesto || false,
        updated_at: new Date().toISOString(),
      }));
      const guardado = filas.length
        ? await supabase.from("movimientos").upsert(filas, { onConflict: "user_id,client_id" })
        : { error: null };
      if (guardado.error) {
        setEstadoNube("Error al guardar; se conserva copia local");
        return;
      }
      const ids = movimientos.map((item) => item.id);
      const borrado = ids.length
        ? await supabase.from("movimientos").delete().eq("user_id", session.user.id).not("client_id", "in", `(${ids.join(",")})`)
        : await supabase.from("movimientos").delete().eq("user_id", session.user.id);
      setEstadoNube(borrado.error ? "Error al sincronizar; se conserva copia local" : "Guardado en la nube");
    }, 500);
    return () => window.clearTimeout(temporizador);
  }, [movimientos, session, nubeLista]);

  useEffect(() => {
    if (session && cargado) {
      localStorage.setItem(
        "mis-finanzas-movimientos",
        JSON.stringify(movimientos),
      );
    }
  }, [movimientos, cargado, session]);

  useEffect(() => {
    if (!session) return;
    queueMicrotask(() => {
      const data = localStorage.getItem("mis-finanzas-presupuestos");
      if (data) {
        try {
          setPresupuestos(JSON.parse(data));
        } catch {
          setPresupuestos({});
        }
      }
      setPresupuestosCargados(true);
    });
  }, [session]);

  useEffect(() => {
    if (session && presupuestosCargados) {
      localStorage.setItem(
        "mis-finanzas-presupuestos",
        JSON.stringify(presupuestos),
      );
    }
  }, [presupuestos, presupuestosCargados, session]);

  function actualizarPresupuesto(categoriaNombre: string, valor: string) {
    const numero = Number(valor);
    setPresupuestos((anteriores) => {
      const copia = { ...anteriores };
      if (!valor || Number.isNaN(numero) || numero <= 0) {
        delete copia[categoriaNombre];
      } else {
        copia[categoriaNombre] = numero;
      }
      return copia;
    });
  }

  const mesesDisponibles = useMemo(
    () => Array.from(new Set(movimientos.map((item) => item.fecha.slice(0, 7)))).sort().reverse(),
    [movimientos],
  );

  const metaPasajes = useMemo(() => {
    const mes = filtroMes === "Todos" ? mesesDisponibles[0] || hoy().slice(0, 7) : filtroMes;
    const [anio, numeroMes] = mes.split("-").map(Number);
    const ultimoDia = new Date(anio, numeroMes, 0).getDate();
    let diasLaborables = 0;
    for (let dia = 1; dia <= ultimoDia; dia += 1) {
      const semana = new Date(anio, numeroMes - 1, dia).getDay();
      if (semana >= 1 && semana <= 5) diasLaborables += 1;
    }
    const movimientosPasajes = movimientos.filter((item) => item.tipo === "gasto" && item.empresa === "Personal" && item.categoria === "Pasajes" && item.fecha.startsWith(mes));
    const gastoComputable = movimientosPasajes.filter((item) => !item.excluirPresupuesto).reduce((total, item) => total + item.monto, 0);
    const pagosEncargo = movimientosPasajes.filter((item) => item.excluirPresupuesto).reduce((total, item) => total + item.monto, 0);
    const porDia = new Map<string, number>();
    movimientosPasajes.filter((item) => !item.excluirPresupuesto).forEach((item) => porDia.set(item.fecha, (porDia.get(item.fecha) || 0) + item.monto));
    const diasSobreMeta = Array.from(porDia.entries()).filter(([, total]) => total > 20).sort(([a], [b]) => b.localeCompare(a));
    const presupuesto = diasLaborables * 20;
    return { mes, diasLaborables, presupuesto, gastoComputable, pagosEncargo, disponible: presupuesto - gastoComputable, porcentaje: presupuesto ? (gastoComputable / presupuesto) * 100 : 0, diasSobreMeta, movimientosPasajes };
  }, [movimientos, filtroMes, mesesDisponibles]);

  const categorias = configuracionLista ? categoriasGastoState : categoriasDefault;
  const categoriasIngreso = configuracionLista ? categoriasIngresoState : categoriasIngresoDefault;
  const subcategoriasExtras = configuracionLista ? subcategoriasExtrasState : subcategoriasExtrasDefault;
  const fuentes = configuracionLista ? fuentesState : fuentesDefault;
  const empresas = configuracionLista ? empresasState : empresasDefault;
  const cuentasPorEmpresa = configuracionLista ? cuentasPorEmpresaState : cuentasPorEmpresaDefault;

  const resumen = useMemo(() => {
    const movimientosMes = filtroMes === "Todos"
      ? movimientos
      : movimientos.filter((item) => item.fecha.startsWith(filtroMes));
    const movimientosPeriodo = filtroEmpresa === "Todas"
      ? movimientosMes
      : movimientosMes.filter((item) =>
          item.tipo === "transferencia"
            ? item.empresaOrigen === filtroEmpresa || item.empresaDestino === filtroEmpresa
            : item.empresa === filtroEmpresa,
        );
    const ingresos = movimientosPeriodo
      .filter((movimiento) => movimiento.tipo === "ingreso" && !["Saldo inicial", "Devolución de préstamo"].includes(movimiento.categoria))
      .reduce((total, movimiento) => total + movimiento.monto, 0);
    const gastos = movimientosPeriodo
      .filter((movimiento) => movimiento.tipo === "gasto" && !esPrestamoJair(movimiento))
      .reduce((total, movimiento) => total + movimiento.monto, 0);
    const porCategoria = categorias
      .map((nombre) => ({
        nombre,
        total: movimientosPeriodo
          .filter(
            (movimiento) =>
              movimiento.tipo === "gasto" &&
              !esPrestamoJair(movimiento) &&
              movimiento.categoria === nombre,
          )
          .reduce((total, movimiento) => total + movimiento.monto, 0),
      }))
      .filter((item) => item.total > 0)
      .sort((a, b) => b.total - a.total);
    const porFuente = fuentes
      .map((nombre) => ({
        nombre,
        total: movimientosPeriodo
          .filter(
            (movimiento) =>
              movimiento.tipo === "gasto" &&
              (movimiento.fuente || "Otro") === nombre,
          )
          .reduce((total, movimiento) => total + movimiento.monto, 0),
      }))
      .filter((item) => item.total > 0)
      .sort((a, b) => b.total - a.total);
    const porExtras = subcategoriasExtras
      .map((nombre) => {
        const movimientosSub = movimientosPeriodo.filter((item) => item.tipo === "gasto" && !esPrestamoJair(item) && item.categoria === "Extras" && (item.subcategoria || "Otros") === nombre);
        const porConcepto = new Map<string, number>();
        movimientosSub.forEach((item) => {
          const clave = item.concepto.trim() || "Sin descripción";
          porConcepto.set(clave, (porConcepto.get(clave) || 0) + item.monto);
        });
        return {
          nombre,
          total: movimientosSub.reduce((total, item) => total + item.monto, 0),
          conceptos: Array.from(porConcepto.entries())
            .map(([concepto, total]) => ({ concepto, total }))
            .sort((a, b) => b.total - a.total),
        };
      })
      .filter((item) => item.total > 0)
      .sort((a, b) => b.total - a.total);
    const porEmpresa = empresas.map((nombre) => {
      const items = movimientosMes.filter((item) => item.tipo !== "transferencia" && item.empresa === nombre);
      const ingresosEmpresa = items.filter((item) => item.tipo === "ingreso" && !["Saldo inicial", "Devolución de préstamo"].includes(item.categoria)).reduce((total, item) => total + item.monto, 0);
      const gastosEmpresa = items.filter((item) => item.tipo === "gasto" && !esPrestamoJair(item)).reduce((total, item) => total + item.monto, 0);
      const movimientosBalance = filtroMes === "Todos" ? movimientos : movimientos.filter((item) => item.fecha <= `${filtroMes}-31`);
      const saldos = new Map<string, number>();
      const sumarCuenta = (cuenta: string | undefined, montoCuenta: number) => {
        const nombreCuenta = cuenta?.trim() || "Otro";
        saldos.set(nombreCuenta, (saldos.get(nombreCuenta) || 0) + montoCuenta);
      };
      movimientosBalance.forEach((item) => {
        if (item.tipo === "ingreso" && (item.empresaDestino || item.empresa) === nombre) sumarCuenta(item.cuentaDestino || cuentaMovimiento(item), item.monto);
        if (item.tipo === "gasto" && (item.empresaOrigen || item.empresa) === nombre) {
          sumarCuenta(item.cuentaOrigen || cuentaMovimiento(item), -item.monto);
          if (nombre === "Personal" && esPrestamoJair(item)) sumarCuenta("Préstamos a Jair", item.monto);
        }
        if (item.tipo === "transferencia" && item.empresaOrigen === nombre) sumarCuenta(item.cuentaOrigen || item.fuente, -item.monto);
        if (item.tipo === "transferencia" && item.empresaDestino === nombre) sumarCuenta(item.cuentaDestino || item.detalleFuente, item.monto);
      });
      const cuentas = Array.from(saldos.entries()).map(([cuenta, saldoCuenta]) => ({ cuenta, saldo: saldoCuenta })).filter((item) => Math.abs(item.saldo) > 0.005).sort((a, b) => b.saldo - a.saldo);
      const balance = cuentas.reduce((total, item) => total + item.saldo, 0);
      return { nombre, ingresos: ingresosEmpresa, gastos: gastosEmpresa, resultado: ingresosEmpresa - gastosEmpresa, balance, cuentas };
    });
    const balance = porEmpresa
      .filter((item) => filtroEmpresa === "Todas" || item.nombre === filtroEmpresa)
      .reduce((total, item) => total + item.balance, 0);
    const personal = porEmpresa.find((item) => item.nombre === "Personal")!;
    const ss = porEmpresa.find((item) => item.nombre === "SS")!;
    const clever = porEmpresa.find((item) => item.nombre === "Clever")!;
    const saldoCuenta = (nombreCuenta: string) => personal.cuentas.find((item) => sinAcentos(item.cuenta) === sinAcentos(nombreCuenta))?.saldo || 0;
    const sip = saldoCuenta("SIP");
    const prestadoJair = saldoCuenta("Préstamos a Jair");
    const disponiblePersonal = personal.balance - sip - prestadoJair;
    const interesesSip = movimientosMes.filter((item) => item.tipo === "ingreso" && item.empresa === "Personal" && item.categoria === "Rendimientos de inversiones" && sinAcentos(item.cuentaDestino || cuentaMovimiento(item)) === "sip").reduce((total, item) => total + item.monto, 0);
    const aportesSip = movimientosMes.filter((item) => (item.tipo === "transferencia" && item.empresaDestino === "Personal" && sinAcentos(item.cuentaDestino || "") === "sip") || (item.tipo === "ingreso" && item.categoria === "Aporte para ahorro" && sinAcentos(item.cuentaDestino || cuentaMovimiento(item)) === "sip")).reduce((total, item) => total + item.monto, 0);
    return { ingresos, gastos, saldo: ingresos - gastos, balance, porCategoria, porFuente, porExtras, porEmpresa, personal, ss, clever, sip, prestadoJair, disponiblePersonal, interesesSip, aportesSip };
  }, [movimientos, filtroMes, filtroEmpresa, categorias, empresas, fuentes, subcategoriasExtras]);

  const mesPresupuesto = filtroMes === "Todos" ? hoy().slice(0, 7) : filtroMes;

  const avancePresupuestos = useMemo(() => {
    const movimientosMes = movimientos.filter(
      (item) => item.tipo === "gasto" && !esPrestamoJair(item) && item.fecha.startsWith(mesPresupuesto),
    );
    return categorias.map((nombre) => {
      const gastado = movimientosMes
        .filter((item) => item.categoria === nombre)
        .reduce((total, item) => total + item.monto, 0);
      const limite = presupuestos[nombre] || 0;
      return {
        nombre,
        gastado,
        limite,
        porcentaje: limite ? (gastado / limite) * 100 : 0,
      };
    });
  }, [movimientos, presupuestos, mesPresupuesto, categorias]);

  const reporteMensual = useMemo(() => {
    const porMes = new Map<string, { ingresos: number; gastos: number }>();
    movimientos.forEach((item) => {
      const mes = item.fecha.slice(0, 7);
      if (!porMes.has(mes)) porMes.set(mes, { ingresos: 0, gastos: 0 });
      const registro = porMes.get(mes)!;
      if (item.tipo === "ingreso" && !["Saldo inicial", "Devolución de préstamo"].includes(item.categoria)) {
        registro.ingresos += item.monto;
      } else if (item.tipo === "gasto" && !esPrestamoJair(item)) {
        registro.gastos += item.monto;
      }
    });
    const meses = Array.from(porMes.entries())
      .map(([mes, valores]) => ({ mes, ...valores, ahorro: valores.ingresos - valores.gastos }))
      .sort((a, b) => a.mes.localeCompare(b.mes))
      .slice(-6);
    const maximo = Math.max(1, ...meses.map((item) => Math.max(item.ingresos, item.gastos)));
    return { meses, maximo };
  }, [movimientos]);

  const lista = movimientos
    .filter(
      (movimiento) =>
        (filtro === "Todos" || movimiento.tipo === filtro.toLowerCase()) &&
        (filtroFuente === "Todas" ||
          (movimiento.fuente || "Otro") === filtroFuente) &&
        (filtroMes === "Todos" || movimiento.fecha.startsWith(filtroMes)) &&
        (filtroEmpresa === "Todas" || (movimiento.tipo === "transferencia"
          ? movimiento.empresaOrigen === filtroEmpresa || movimiento.empresaDestino === filtroEmpresa
          : movimiento.empresa === filtroEmpresa)) &&
        (!soloJair || esMovimientoJair(movimiento)) &&
        movimiento.concepto.toLowerCase().includes(busqueda.toLowerCase()),
    )
    .sort((a, b) => b.fecha.localeCompare(a.fecha));

  const actual = borradores[indiceRevision];
  const archivoActivo = archivosExcel[indiceArchivo];
  const totalFilasExcel = archivosExcel.reduce(
    (total, archivo) => total + archivo.filas.length,
    0,
  );
  const totalAImportar = borradores.filter(
    (movimiento) => !movimiento.descartado,
  ).length;

  function guardar(e: FormEvent) {
    e.preventDefault();
    const valor = Number(monto);
    if (!valor) return;
    if (tipo === "transferencia" && empresaOrigen === empresaDestino && cuentaOrigen === cuentaDestino) {
      alert("La cuenta de origen y destino deben ser diferentes.");
      return;
    }
    const detalleCalculado = detalleAutomatico(fuente);
    const cuentaSeleccionada = detalleCalculado || detalleFuente.trim() || fuente;
    const esDevolucionJair = tipo === "ingreso" && categoriaIngreso === "Devolución de préstamo";
    const movimientoGuardado: Movimiento = {
        id: movimientoEditando ?? Date.now(),
        tipo: esDevolucionJair ? "transferencia" : tipo,
        monto: valor,
        concepto: concepto.trim() || (esDevolucionJair ? "Devolución de préstamo de Jair" : tipo === "transferencia" ? "Transferencia propia" : "Sin concepto"),
        categoria: esDevolucionJair ? "Devolución de préstamo" : tipo === "ingreso" ? categoriaIngreso : tipo === "transferencia" ? "Transferencias entre cuentas" : categoria,
        fuente: esDevolucionJair ? "Préstamos a Jair" : tipo === "transferencia" ? cuentaOrigen : fuente,
        detalleFuente: esDevolucionJair ? cuentaSeleccionada : tipo === "transferencia" ? cuentaDestino : detalleCalculado || detalleFuente.trim(),
        fecha,
        empresa: esDevolucionJair ? "Personal" : tipo === "transferencia" ? empresaOrigen : empresa,
        subcategoria: tipo === "gasto" && categoria === "Extras" ? subcategoria : "",
        empresaOrigen: esDevolucionJair ? "Personal" : tipo === "ingreso" ? undefined : tipo === "transferencia" ? empresaOrigen : empresa,
        empresaDestino: esDevolucionJair ? "Personal" : tipo === "gasto" ? undefined : tipo === "transferencia" ? empresaDestino : empresa,
        cuentaOrigen: esDevolucionJair ? "Préstamos a Jair" : tipo === "ingreso" ? undefined : tipo === "transferencia" ? cuentaOrigen : cuentaSeleccionada,
        cuentaDestino: esDevolucionJair ? cuentaSeleccionada : tipo === "gasto" ? undefined : tipo === "transferencia" ? cuentaDestino : cuentaSeleccionada,
        claseTransferencia: esDevolucionJair ? "Devolución de préstamo" : tipo === "transferencia" ? claseTransferencia : undefined,
        excluirPresupuesto: tipo === "gasto" && categoria === "Pasajes" ? excluirPresupuesto : false,
      };
    setMovimientos((anteriores) => movimientoEditando === null
      ? [movimientoGuardado, ...anteriores]
      : anteriores.map((movimiento) => movimiento.id === movimientoEditando ? movimientoGuardado : movimiento));
    setMonto("");
    setConcepto("");
    setDetalleFuente("");
    setCuentaOrigen("BCP");
    setCuentaDestino("IBK");
    setEmpresaOrigen("Personal");
    setEmpresaDestino("Personal");
    setClaseTransferencia("Entre cuentas propias");
    setExcluirPresupuesto(false);
    setSubcategoria("Otros");
    setCategoriaIngreso("Sueldo");
    setFecha(hoy());
    setMovimientoEditando(null);
    setModal(false);
  }

  function editar(movimiento: Movimiento) {
    setMovimientoEditando(movimiento.id);
    setTipo(movimiento.tipo);
    setMonto(String(movimiento.monto));
    setConcepto(movimiento.concepto === "Sin concepto" ? "" : movimiento.concepto);
    setCategoria(movimiento.tipo === "gasto" ? movimiento.categoria : categorias[0] || "Comida");
    setCategoriaIngreso(movimiento.tipo === "ingreso" ? movimiento.categoria : "Sueldo");
    setEmpresa(movimiento.empresa || "Personal");
    setSubcategoria(movimiento.subcategoria || "Otros");
    setFuente(movimiento.tipo === "transferencia" ? "Transferencia" : movimiento.fuente || "Otro");
    setDetalleFuente(movimiento.tipo === "transferencia" ? "" : movimiento.detalleFuente || "");
    setEmpresaOrigen(movimiento.empresaOrigen || movimiento.empresa || "Personal");
    setEmpresaDestino(movimiento.empresaDestino || movimiento.empresa || "Personal");
    setCuentaOrigen(movimiento.cuentaOrigen || movimiento.fuente || "BCP");
    setCuentaDestino(movimiento.cuentaDestino || movimiento.detalleFuente || "IBK");
    setClaseTransferencia(movimiento.claseTransferencia || "Entre cuentas propias");
    setExcluirPresupuesto(Boolean(movimiento.excluirPresupuesto));
    setFecha(movimiento.fecha);
    setModal(true);
  }

  function cerrarModalMovimiento() {
    setMovimientoEditando(null);
    setModal(false);
  }

  function borrar(id: number) {
    if (confirm("¿Eliminar este movimiento?")) {
      setMovimientos((anteriores) =>
        anteriores.filter((movimiento) => movimiento.id !== id),
      );
    }
  }

  function verDetalle(tipoDetalle: "Todos" | "Ingreso" | "Gasto" | "Transferencia", empresaDetalle?: Empresa | "Todas") {
    setSoloJair(false);
    setFiltro(tipoDetalle);
    if (empresaDetalle) setFiltroEmpresa(empresaDetalle);
    setVista("movimientos");
  }

  function verDetalleJair() {
    setSoloJair(true);
    setFiltro("Todos");
    setFiltroEmpresa("Personal");
    setFiltroFuente("Todas");
    setBusqueda("");
    setVista("movimientos");
  }

  function alternarPresupuesto(id: number) {
    setMovimientos((anteriores) => anteriores.map((movimiento) => movimiento.id === id ? { ...movimiento, excluirPresupuesto: !movimiento.excluirPresupuesto } : movimiento));
  }

  async function accederCuenta(accion: "ingresar" | "crear") {
    if (!email.trim() || password.length < 6) {
      setMensajeCuenta("Ingresa un correo válido y una contraseña de al menos 6 caracteres.");
      return;
    }
    setProcesandoCuenta(true);
    setMensajeCuenta("");
    const { error } = accion === "crear"
      ? await supabase.auth.signUp({
          email: email.trim(),
          password,
          options: { emailRedirectTo: window.location.origin },
        })
      : await supabase.auth.signInWithPassword({ email: email.trim(), password });
    setProcesandoCuenta(false);
    if (error) {
      setMensajeCuenta(error.message);
    } else if (accion === "crear") {
      setMensajeCuenta("Cuenta creada. Revisa tu correo para confirmar el acceso.");
    } else {
      setModalCuenta(false);
      setPassword("");
    }
  }

  async function cerrarSesion() {
    await supabase.auth.signOut();
    setModalCuenta(false);
    setPassword("");
  }

  function reiniciarImportacion() {
    setArchivosExcel([]);
    setIndiceArchivo(0);
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

  async function procesarArchivo(archivo: File): Promise<ArchivoImportado> {
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
    return {
      id: `${archivo.name}-${archivo.size}-${archivo.lastModified}`,
      nombre: archivo.name,
      filas,
      columnas,
      mapa: detectarMapa(columnas),
      positivosSon: "gasto",
    };
  }

  async function cargarExcel(e: ChangeEvent<HTMLInputElement>) {
    const seleccionados = Array.from(e.target.files || []);
    if (!seleccionados.length) return;
    setErrorImportacion("");

    const idsExistentes = new Set(archivosExcel.map((archivo) => archivo.id));
    const nuevos: ArchivoImportado[] = [];
    const errores: string[] = [];
    for (const archivo of seleccionados) {
      const id = `${archivo.name}-${archivo.size}-${archivo.lastModified}`;
      if (idsExistentes.has(id)) continue;
      try {
        const procesado = await procesarArchivo(archivo);
        nuevos.push(procesado);
        idsExistentes.add(id);
      } catch {
        errores.push(archivo.name);
      }
    }

    if (nuevos.length) {
      setIndiceArchivo(archivosExcel.length);
      setArchivosExcel((anteriores) => [...anteriores, ...nuevos]);
    }
    if (errores.length) {
      setErrorImportacion(
        `No pudimos leer ${errores.join(", ")}. Verifica que cada archivo sea .xlsx o .csv y tenga encabezados en la primera fila.`,
      );
    } else if (!nuevos.length) {
      setErrorImportacion("Los archivos seleccionados ya fueron agregados.");
    }
    e.target.value = "";
  }

  function actualizarMapaActivo(campo: CampoImportacion, valor: string) {
    if (!archivoActivo) return;
    setArchivosExcel((anteriores) =>
      anteriores.map((archivo, indice) =>
        indice === indiceArchivo
          ? { ...archivo, mapa: { ...archivo.mapa, [campo]: valor } }
          : archivo,
      ),
    );
  }

  function actualizarPositivosActivo(valor: Tipo) {
    setArchivosExcel((anteriores) =>
      anteriores.map((archivo, indice) =>
        indice === indiceArchivo
          ? { ...archivo, positivosSon: valor }
          : archivo,
      ),
    );
  }

  function quitarArchivo(id: string) {
    const posicion = archivosExcel.findIndex((archivo) => archivo.id === id);
    const restantes = archivosExcel.filter((archivo) => archivo.id !== id);
    setArchivosExcel(restantes);
    if (!restantes.length) {
      setIndiceArchivo(0);
    } else if (posicion < indiceArchivo) {
      setIndiceArchivo(indiceArchivo - 1);
    } else if (indiceArchivo >= restantes.length) {
      setIndiceArchivo(restantes.length - 1);
    }
  }

  function valorFila(fila: FilaExcel, columna: string) {
    return columna ? fila[columna] : "";
  }

  function prepararRevision() {
    setErrorImportacion("");
    const indiceIncompleto = archivosExcel.findIndex(
      (archivo) =>
        !archivo.mapa.fecha ||
        !archivo.mapa.concepto ||
        (!archivo.mapa.monto &&
          !archivo.mapa.cargo &&
          !archivo.mapa.abono),
    );
    if (indiceIncompleto >= 0) {
      setIndiceArchivo(indiceIncompleto);
      setErrorImportacion(
        `Completa las columnas requeridas de ${archivosExcel[indiceIncompleto].nombre}: fecha, descripción y monto o cargos/abonos.`,
      );
      return;
    }

    const existentes = new Set(movimientos.map(claveMovimiento));
    const nuevasClaves = new Set<string>();
    let duplicados = 0;
    let secuencia = 0;
    const baseId = Date.now();
    const importados: MovimientoImportado[] = [];
    for (const archivo of archivosExcel) {
      const mapa = archivo.mapa;
      for (const fila of archivo.filas) {
        secuencia += 1;
        const conceptoImportado = String(
          valorFila(fila, mapa.concepto) ?? "",
        ).trim();
        const fechaImportada = fechaDesdeExcel(valorFila(fila, mapa.fecha));
        if (!conceptoImportado || !fechaImportada) continue;

        const cargo = numeroDesdeExcel(valorFila(fila, mapa.cargo));
        const abono = numeroDesdeExcel(valorFila(fila, mapa.abono));
        const montoUnico = numeroDesdeExcel(valorFila(fila, mapa.monto));
        let tipoImportado =
          tipoDesdeTexto(valorFila(fila, mapa.tipo)) || archivo.positivosSon;
        let montoImportado = Math.abs(montoUnico);

        if (mapa.cargo || mapa.abono) {
          if (Math.abs(cargo) > 0) {
            tipoImportado = "gasto";
            montoImportado = Math.abs(cargo);
          } else if (Math.abs(abono) > 0) {
            tipoImportado = "ingreso";
            montoImportado = Math.abs(abono);
          }
        } else if (!mapa.tipo && montoUnico < 0) {
          tipoImportado = "gasto";
        }
        if (!montoImportado) continue;

        const fuenteImportada = normalizarFuente(
          valorFila(fila, mapa.fuente),
        );
        const movimiento: MovimientoImportado = {
          id: baseId + secuencia,
          tipo: tipoImportado,
          monto: montoImportado,
          concepto: conceptoImportado,
          categoria:
            tipoImportado === "ingreso"
              ? categoriasIngreso[0]
              : sugerirCategoria(conceptoImportado),
          fuente: fuenteImportada.fuente,
          detalleFuente: fuenteImportada.detalle,
          fecha: fechaImportada,
          empresa: "Personal",
          subcategoria: "",
        };
        const clave = claveMovimiento(movimiento);
        if (existentes.has(clave) || nuevasClaves.has(clave)) {
          duplicados += 1;
          continue;
        }
        nuevasClaves.add(clave);
        importados.push(movimiento);
      }
    }

    if (!importados.length) {
      setErrorImportacion(
        duplicados
          ? "Todos los movimientos de los archivos ya están registrados."
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
          ? categoriasIngreso[0]
          : categoriasIngreso.includes(actual?.categoria || "")
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
          empresa: empresaMovimiento,
          subcategoria: subcategoriaMovimiento,
        }): Movimiento => ({
          id,
          tipo: tipoMovimiento,
          monto: montoMovimiento,
          concepto: conceptoMovimiento,
          categoria: categoriaMovimiento,
          fuente: fuenteMovimiento,
          detalleFuente: detalleFuenteMovimiento,
          fecha: fechaMovimiento,
          empresa: empresaMovimiento || "Personal",
          subcategoria: subcategoriaMovimiento || "",
        }),
      );
    if (confirmados.length) {
      setMovimientos((anteriores) => [...confirmados, ...anteriores]);
    }
    setModalRevision(false);
    reiniciarImportacion();
  }

  if (!authRevisada) {
    return (
      <main className="login-shell login-loading" aria-live="polite">
        <div className="login-brand"><span>MF</span><strong>Mis Finanzas</strong></div>
        <p>Validando acceso seguro…</p>
      </main>
    );
  }

  if (!session) {
    return (
      <main className="login-shell">
        <section className="login-panel">
          <div className="login-brand"><span>MF</span><strong>Mis Finanzas</strong></div>
          <div className="login-copy">
            <span className="eyebrow">ACCESO SEGURO</span>
            <h1>Controla tus finanzas desde cualquier dispositivo.</h1>
            <p>Inicia sesión para consultar, registrar e importar tus movimientos. Tu información permanece privada y sincronizada.</p>
          </div>
          <form className="login-form" onSubmit={(e) => { e.preventDefault(); accederCuenta("ingresar"); }}>
            <label>Correo electrónico<input autoFocus type="email" autoComplete="email" placeholder="nombre@correo.com" value={email} onChange={(e) => setEmail(e.target.value)} required /></label>
            <label>Contraseña<input type="password" autoComplete="current-password" minLength={6} placeholder="Mínimo 6 caracteres" value={password} onChange={(e) => setPassword(e.target.value)} required /></label>
            {mensajeCuenta && <p className="account-message">{mensajeCuenta}</p>}
            <button className="save" disabled={procesandoCuenta}>{procesandoCuenta ? "Procesando…" : "Iniciar sesión"}</button>
            <button type="button" className="create-account" disabled={procesandoCuenta} onClick={() => accederCuenta("crear")}>Crear una cuenta</button>
          </form>
          <small className="login-security">Acceso protegido mediante Supabase</small>
        </section>
        <aside className="login-visual" aria-hidden="true"><div><span>Saldo, gastos e ingresos</span><strong>Todo en un solo lugar.</strong><i /></div></aside>
      </main>
    );
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span>MF</span>
          <strong>Mis Finanzas</strong>
        </div>
        <nav>
          <button className={vista === "inicio" ? "active" : ""} onClick={() => setVista("inicio")}>
            ⌂ <span>Inicio</span>
          </button>
          <button className={vista === "movimientos" ? "active" : ""} onClick={() => setVista("movimientos")}>
            ⇄ <span>Movimientos</span>
          </button>
          <button className={vista === "presupuestos" ? "active" : ""} onClick={() => setVista("presupuestos")}>
            ◔ <span>Presupuestos</span>
          </button>
          <button className={vista === "metas" ? "active" : ""} onClick={() => setVista("metas")}>
            ◎ <span>Metas</span>
          </button>
          <button className={vista === "reportes" ? "active" : ""} onClick={() => setVista("reportes")}>
            ▥ <span>Reportes</span>
          </button>
          <button className={vista === "ajustes" ? "active" : ""} onClick={() => setVista("ajustes")}>
            ⚙ <span>Ajustes</span>
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
            <h1>{vista === "inicio" ? "Resumen de patrimonio" : vista === "metas" ? "Metas financieras" : vista === "presupuestos" ? "Presupuestos por categoría" : vista === "reportes" ? "Reportes" : vista === "ajustes" ? "Ajustes" : "Movimientos por unidad"}</h1>
            <p className="subtitle">
              {vista === "inicio"
                ? "Personal, ahorro SIP y negocios claramente separados."
                : vista === "metas" ? "Control mensual para tomar decisiones antes de exceder el presupuesto."
                : vista === "presupuestos" ? "Define un límite mensual por categoría y sigue tu avance."
                : vista === "reportes" ? "La evolución de tus ingresos y gastos a lo largo del tiempo."
                : vista === "ajustes" ? "Personaliza tus categorías, medios de pago y unidades de negocio."
                : "Consulta el detalle cuando lo necesites, sin recargar el panel principal."}
            </p>
          </div>
          <div className="account-area">
            <span className={session ? "cloud-state online" : "cloud-state"}>{estadoNube}</span>
            <button className="avatar" aria-label="Cuenta y sincronización" onClick={() => setModalCuenta(true)}>
              {session?.user.email?.slice(0, 2).toUpperCase() || "RS"}
            </button>
          </div>
        </header>

        {vista !== "ajustes" && <div className="scope-bar card">
          {vista !== "metas" && vista !== "presupuestos" && vista !== "reportes" && <div><span>Unidad</span><select value={filtroEmpresa} onChange={(e) => setFiltroEmpresa(e.target.value)}><option>Todas</option>{empresas.map((item) => <option key={item}>{item}</option>)}</select></div>}
          <div><span>Periodo</span><select value={filtroMes} onChange={(e) => setFiltroMes(e.target.value)}><option value="Todos">Todos los meses</option>{mesesDisponibles.map((mes) => <option key={mes} value={mes}>{new Date(`${mes}-01T12:00:00`).toLocaleDateString("es-PE", { month: "long", year: "numeric" })}</option>)}</select></div>
          <button onClick={() => { setMovimientoEditando(null); setModal(true); }}>＋ Nuevo movimiento</button>
        </div>}

        {vista === "inicio" && <section className="financial-summary">
          <article className="card summary-unit personal-unit" onClick={() => setFiltroEmpresa("Personal")}>
            <div><span className="unit-icon">P</span><span className="eyebrow">FINANZAS PERSONALES</span></div>
            <h2>Personal</h2>
            <strong>{ocultar ? "S/ ••••••" : soles(resumen.disponiblePersonal)}</strong>
            <p>Dinero personal disponible, sin SIP ni préstamos por cobrar.</p>
            <section>
              <span className="detail-link" onClick={(e) => { e.stopPropagation(); verDetalle("Ingreso", "Personal"); }}>Ingresos del periodo <b>{soles(resumen.personal.ingresos)}</b></span>
              <span className="detail-link" onClick={(e) => { e.stopPropagation(); verDetalle("Gasto", "Personal"); }}>Gastos reales <b>{soles(resumen.personal.gastos)}</b></span>
              <span className="detail-link" onClick={(e) => { e.stopPropagation(); verDetalleJair(); }}>Prestado a Jair <b>{soles(resumen.prestadoJair)}</b></span>
            </section>
          </article>
          <article className="card summary-unit sip-unit">
            <div><span className="unit-icon">S</span><span className="eyebrow">AHORRO PARA DEPARTAMENTO</span></div>
            <h2>SIP</h2>
            <strong>{ocultar ? "S/ ••••••" : soles(resumen.sip)}</strong>
            <p>Meta {soles(40000)} · {Math.min(100, Math.max(0, (resumen.sip / 40000) * 100)).toFixed(1)}% alcanzado</p>
            <div className="goal-progress"><i style={{ width: `${Math.min(100, Math.max(0, (resumen.sip / 40000) * 100))}%` }} /></div>
            <section>
              <span className="detail-link" onClick={() => verDetalle("Transferencia", "Personal")}>Depósitos del periodo <b>{soles(resumen.aportesSip)}</b></span>
              <span className="detail-link" onClick={() => verDetalle("Ingreso", "Personal")}>Intereses del periodo <b>{soles(resumen.interesesSip)}</b></span>
              <span>Falta para la meta <b>{soles(Math.max(0, 40000 - resumen.sip))}</b></span>
            </section>
          </article>
          {[resumen.ss, resumen.clever].map((unidad) => <article className="card summary-unit business-unit" key={unidad.nombre} onClick={() => setFiltroEmpresa(unidad.nombre)}>
            <div><span className="unit-icon">{unidad.nombre.slice(0, 1)}</span><span className="eyebrow">NEGOCIO INDEPENDIENTE</span></div>
            <h2>{unidad.nombre}</h2>
            <strong>{ocultar ? "S/ ••••••" : soles(unidad.balance)}</strong>
            <p>Cuenta y operación separada de sus finanzas personales.</p>
            <section>
              <span className="detail-link" onClick={(e) => { e.stopPropagation(); verDetalle("Ingreso", unidad.nombre as Empresa); }}>Ventas <b>{soles(unidad.ingresos)}</b></span>
              <span className="detail-link" onClick={(e) => { e.stopPropagation(); verDetalle("Gasto", unidad.nombre as Empresa); }}>Compras y egresos <b>{soles(unidad.gastos)}</b></span>
              <span>Resultado del periodo <b className={unidad.resultado >= 0 ? "ingreso" : "gasto"}>{soles(unidad.resultado)}</b></span>
            </section>
          </article>)}
          <button className="privacy-toggle" onClick={() => setOcultar(!ocultar)}>{ocultar ? "Mostrar importes" : "Ocultar importes"}</button>
        </section>}

        <section className={vista === "inicio" ? "lower-grid" : "lower-grid movements-view"}>
          {vista === "inicio" && <>
          <article className="card distribution">
            <div className="section-title">
              <div>
                <span className="eyebrow">ANÁLISIS</span>
                <h2>¿En qué estás gastando?</h2>
              </div>
              <span className="period">
                {filtroMes === "Todos"
                  ? "Todos"
                  : new Date(`${filtroMes}-01T12:00:00`).toLocaleDateString("es-PE", { month: "long", year: "numeric" })}
              </span>
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
                          `${colorCategoria(item.nombre, indice)} ${
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
                  {resumen.porCategoria.map((item, indice) => (
                    <div key={item.nombre}>
                      <i style={{ background: colorCategoria(item.nombre, indice) }} />
                      <span>{item.nombre}</span>
                      <b>{Math.round((item.total / resumen.gastos) * 100)}%</b>
                      <small>{soles(item.total)}</small>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </article>

          <article className="card financial-lens">
            <div className="section-title"><div><span className="eyebrow">CONTROL PERSONAL</span><h2>Dinero con destino específico</h2></div></div>
            <div className="financial-focus">
              <div><span>Préstamos pendientes de Jair</span><strong>{soles(resumen.prestadoJair)}</strong><small>No se consideran gasto: permanecen como dinero por cobrar.</small></div>
              <div><span>Ahorro acumulado en SIP</span><strong>{soles(resumen.sip)}</strong><small>Incluye depósitos e intereses registrados en la cuenta SIP.</small></div>
            </div>
            <div className="extras-detail">
              <span className="eyebrow">GASTOS EXTRAS REALES</span>
              {resumen.porExtras.length ? resumen.porExtras.map((item) => (
                <div className="extras-group" key={item.nombre}>
                  <div className="extras-group-head" onClick={() => setExtrasAbierto(extrasAbierto === item.nombre ? null : item.nombre)}>
                    <span>{item.nombre}</span>
                    <b>{soles(item.total)}</b>
                  </div>
                  {extrasAbierto === item.nombre && (
                    <div className="extras-conceptos">
                      {item.conceptos.map((concepto) => (
                        <div key={concepto.concepto}><span>{concepto.concepto}</span><b>{soles(concepto.total)}</b></div>
                      ))}
                    </div>
                  )}
                </div>
              )) : <p>No hay otros gastos clasificados como Extras en este periodo.</p>}
            </div>
          </article>
          </>}

          {vista === "metas" && <article className="card goals-panel">
            <div className="section-title"><div><span className="eyebrow">MOVILIDAD PERSONAL</span><h2>Presupuesto de pasajes</h2></div><span className="period">{new Date(`${metaPasajes.mes}-01T12:00:00`).toLocaleDateString("es-PE", { month: "long", year: "numeric" })}</span></div>
            <div className="goal-overview">
              <div><span>Presupuesto mensual</span><strong>{soles(metaPasajes.presupuesto)}</strong><small>S/20 × {metaPasajes.diasLaborables} días laborables</small></div>
              <div><span>Gasto computable</span><strong className={metaPasajes.disponible >= 0 ? "ingreso" : "gasto"}>{soles(metaPasajes.gastoComputable)}</strong><small>{metaPasajes.porcentaje.toFixed(1)}% utilizado</small></div>
              <div><span>{metaPasajes.disponible >= 0 ? "Disponible" : "Exceso"}</span><strong className={metaPasajes.disponible >= 0 ? "ingreso" : "gasto"}>{soles(Math.abs(metaPasajes.disponible))}</strong><small>Pagos por encargo excluidos: {soles(metaPasajes.pagosEncargo)}</small></div>
            </div>
            <div className="monthly-progress"><i className={metaPasajes.porcentaje > 100 ? "over" : ""} style={{ width: `${Math.min(100, metaPasajes.porcentaje)}%` }} /></div>
            <p className="goal-rule">La referencia diaria es S/20, pero la evaluación principal se realiza contra el presupuesto total del mes. Los pagos hechos por encargo se conservan en el historial y no consumen esta meta.</p>
            <div className="goal-columns">
              <section><h3>Días sobre S/20</h3>{metaPasajes.diasSobreMeta.length ? metaPasajes.diasSobreMeta.map(([dia, total]) => <div key={dia}><span>{new Date(`${dia}T12:00:00`).toLocaleDateString("es-PE", { day: "2-digit", month: "short" })}</span><b>{soles(total)}</b></div>) : <p>No hay días por encima de la referencia.</p>}</section>
              <section><h3>Movimientos de pasajes</h3>{metaPasajes.movimientosPasajes.length ? metaPasajes.movimientosPasajes.sort((a, b) => b.fecha.localeCompare(a.fecha)).map((movimiento) => <div key={movimiento.id} className={movimiento.excluirPresupuesto ? "excluded" : ""}><span>{new Date(`${movimiento.fecha}T12:00:00`).toLocaleDateString("es-PE", { day: "2-digit", month: "short" })} · {movimiento.concepto}</span><b>{soles(movimiento.monto)}</b><button onClick={() => alternarPresupuesto(movimiento.id)}>{movimiento.excluirPresupuesto ? "Incluir en meta" : "Pago por encargo"}</button></div>) : <p>No hay pasajes registrados en este mes.</p>}</section>
            </div>
          </article>}

          {vista === "presupuestos" && <article className="card budgets-panel">
            <div className="section-title">
              <div>
                <span className="eyebrow">GASTOS PERSONALES</span>
                <h2>Presupuesto de {new Date(`${mesPresupuesto}-01T12:00:00`).toLocaleDateString("es-PE", { month: "long", year: "numeric" })}</h2>
              </div>
            </div>
            <p className="budgets-hint">Define cuánto quieres gastar como máximo en cada categoría este mes. Se guarda en este dispositivo; déjalo en 0 para no controlar esa categoría.</p>
            <div className="budgets-grid">
              {avancePresupuestos.map((item) => (
                <div className="budget-card" key={item.nombre}>
                  <div className="budget-card-head">
                    <span><i style={{ background: colores[item.nombre] || "#94a3b8" }} />{item.nombre}</span>
                    <button className="detail-link-btn" onClick={() => verDetalle("Gasto")}>Ver movimientos</button>
                  </div>
                  <div className="budget-numbers">
                    <b className={item.limite && item.gastado > item.limite ? "gasto" : ""}>{soles(item.gastado)}</b>
                    <span>de</span>
                    <input
                      type="number"
                      min={0}
                      placeholder="Sin límite"
                      value={presupuestos[item.nombre] ?? ""}
                      onChange={(e) => actualizarPresupuesto(item.nombre, e.target.value)}
                    />
                  </div>
                  {item.limite > 0 && (
                    <>
                      <div className="monthly-progress"><i className={item.porcentaje > 100 ? "over" : ""} style={{ width: `${Math.min(100, item.porcentaje)}%` }} /></div>
                      <small className={item.porcentaje > 100 ? "gasto" : ""}>
                        {item.porcentaje > 100
                          ? `Superaste el límite en ${soles(item.gastado - item.limite)}`
                          : `${soles(item.limite - item.gastado)} disponibles (${item.porcentaje.toFixed(0)}%)`}
                      </small>
                    </>
                  )}
                </div>
              ))}
            </div>
          </article>}

          {vista === "reportes" && <article className="card reports-panel">
            <div className="section-title">
              <div>
                <span className="eyebrow">TENDENCIA</span>
                <h2>Ingresos vs. gastos por mes</h2>
              </div>
            </div>
            {reporteMensual.meses.length === 0 ? (
              <p className="empty">Aún no hay suficientes movimientos para mostrar una tendencia.</p>
            ) : (
              <>
                <div className="trend-chart">
                  {reporteMensual.meses.map((item) => (
                    <div className="trend-column" key={item.mes}>
                      <div className="trend-bars">
                        <i className="bar-ingreso" style={{ height: `${(item.ingresos / reporteMensual.maximo) * 100}%` }} title={`Ingresos: ${soles(item.ingresos)}`} />
                        <i className="bar-gasto" style={{ height: `${(item.gastos / reporteMensual.maximo) * 100}%` }} title={`Gastos: ${soles(item.gastos)}`} />
                      </div>
                      <span>{new Date(`${item.mes}-01T12:00:00`).toLocaleDateString("es-PE", { month: "short" })}</span>
                    </div>
                  ))}
                </div>
                <div className="trend-legend">
                  <span><i className="bar-ingreso" /> Ingresos</span>
                  <span><i className="bar-gasto" /> Gastos</span>
                </div>
                <div className="trend-table">
                  {reporteMensual.meses.slice().reverse().map((item) => (
                    <div key={item.mes}>
                      <span>{new Date(`${item.mes}-01T12:00:00`).toLocaleDateString("es-PE", { month: "long", year: "numeric" })}</span>
                      <b className="ingreso">{soles(item.ingresos)}</b>
                      <b className="gasto">{soles(item.gastos)}</b>
                      <b className={item.ahorro >= 0 ? "ingreso" : "gasto"}>{soles(item.ahorro)}</b>
                    </div>
                  ))}
                </div>
              </>
            )}
          </article>}

          {vista === "ajustes" && <article className="card settings-panel">
            {!session ? (
              <div className="settings-locked">
                <p>Inicia sesión o crea una cuenta para personalizar tus propias categorías, medios de pago y unidades de negocio. Mientras tanto, la app usa una lista básica compartida.</p>
                <button className="save" style={{ width: "auto", marginTop: 4 }} onClick={() => setModalCuenta(true)}>Iniciar sesión o crear cuenta</button>
              </div>
            ) : (
              <div className="settings-grid">
                <div className="settings-group">
                  <h3>Categorías de gasto</h3>
                  <div className="chip-list">
                    {categoriasGastoState.map((item) => (
                      <span className="chip" key={item}>{item}<button onClick={() => eliminarCategoriaGasto(item)} aria-label={`Quitar ${item}`}>×</button></span>
                    ))}
                    {!categoriasGastoState.length && <p className="settings-empty">Aún no tienes categorías de gasto.</p>}
                  </div>
                  <div className="settings-add">
                    <input placeholder="Nueva categoría" value={nuevaCategoria} onChange={(e) => setNuevaCategoria(e.target.value)} onKeyDown={(e) => e.key === "Enter" && agregarCategoriaGasto()} />
                    <button onClick={agregarCategoriaGasto}>Agregar</button>
                  </div>
                </div>

                <div className="settings-group">
                  <h3>Categorías de ingreso</h3>
                  <div className="chip-list">
                    {categoriasIngresoState.map((item) => (
                      <span className="chip" key={item}>{item}<button onClick={() => eliminarCategoriaIngresoNueva(item)} aria-label={`Quitar ${item}`}>×</button></span>
                    ))}
                    {!categoriasIngresoState.length && <p className="settings-empty">Aún no tienes categorías de ingreso.</p>}
                  </div>
                  <div className="settings-add">
                    <input placeholder="Nueva categoría" value={nuevaCategoriaIngreso} onChange={(e) => setNuevaCategoriaIngreso(e.target.value)} onKeyDown={(e) => e.key === "Enter" && agregarCategoriaIngresoNueva()} />
                    <button onClick={agregarCategoriaIngresoNueva}>Agregar</button>
                  </div>
                </div>

                <div className="settings-group">
                  <h3>Subcategorías de &quot;Extras&quot;</h3>
                  <div className="chip-list">
                    {subcategoriasExtrasState.map((item) => (
                      <span className="chip" key={item}>{item}<button onClick={() => eliminarSubcategoriaNueva(item)} aria-label={`Quitar ${item}`}>×</button></span>
                    ))}
                    {!subcategoriasExtrasState.length && <p className="settings-empty">Aún no tienes subcategorías.</p>}
                  </div>
                  <div className="settings-add">
                    <input placeholder="Nueva subcategoría" value={nuevaSubcategoria} onChange={(e) => setNuevaSubcategoria(e.target.value)} onKeyDown={(e) => e.key === "Enter" && agregarSubcategoriaNueva()} />
                    <button onClick={agregarSubcategoriaNueva}>Agregar</button>
                  </div>
                </div>

                <div className="settings-group">
                  <h3>Medios de pago</h3>
                  <div className="chip-list">
                    {fuentesState.map((item) => (
                      <span className="chip" key={item}>{item}<button onClick={() => eliminarFuenteNueva(item)} aria-label={`Quitar ${item}`}>×</button></span>
                    ))}
                    {!fuentesState.length && <p className="settings-empty">Aún no tienes medios de pago.</p>}
                  </div>
                  <div className="settings-add">
                    <input placeholder="Nuevo medio de pago" value={nuevaFuente} onChange={(e) => setNuevaFuente(e.target.value)} onKeyDown={(e) => e.key === "Enter" && agregarFuenteNueva()} />
                    <button onClick={agregarFuenteNueva}>Agregar</button>
                  </div>
                </div>

                <div className="settings-group settings-group-wide">
                  <h3>Unidades de negocio</h3>
                  <p className="settings-hint">Estas son solo tuyas. Cada persona que use la app crea las suyas propias.</p>
                  <div className="chip-list">
                    {empresasState.map((item) => (
                      <span className={`chip ${empresaAjustes === item ? "selected" : ""}`} key={item} onClick={() => setEmpresaAjustes(item)}>{item}<button onClick={(e) => { e.stopPropagation(); eliminarEmpresaNueva(item); }} aria-label={`Quitar ${item}`}>×</button></span>
                    ))}
                    {!empresasState.length && <p className="settings-empty">Aún no tienes unidades de negocio.</p>}
                  </div>
                  <div className="settings-add">
                    <input placeholder="Nueva unidad (ej. Personal)" value={nuevaEmpresa} onChange={(e) => setNuevaEmpresa(e.target.value)} onKeyDown={(e) => e.key === "Enter" && agregarEmpresaNueva()} />
                    <button onClick={agregarEmpresaNueva}>Agregar</button>
                  </div>

                  {empresasState.length > 0 && (
                    <div className="settings-subsection">
                      <h4>Cuentas de {empresaAjustes}</h4>
                      <div className="chip-list">
                        {(cuentasPorEmpresaState[empresaAjustes] || []).map((item) => (
                          <span className="chip" key={item}>{item}<button onClick={() => eliminarCuentaNueva(item)} aria-label={`Quitar ${item}`}>×</button></span>
                        ))}
                        {!(cuentasPorEmpresaState[empresaAjustes] || []).length && <p className="settings-empty">Sin cuentas todavía.</p>}
                      </div>
                      <div className="settings-add">
                        <input placeholder="Nueva cuenta" value={nuevaCuenta} onChange={(e) => setNuevaCuenta(e.target.value)} onKeyDown={(e) => e.key === "Enter" && agregarCuentaNueva()} />
                        <button onClick={agregarCuentaNueva}>Agregar</button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </article>}

          {vista === "movimientos" && <article className="card movements">
            <div className="section-title">
              <div>
                <span className="eyebrow">ACTIVIDAD</span>
                <h2>Historial de movimientos</h2>
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
                  <option>Transferencia</option>
                </select>
                <select aria-label="Mes" value={filtroMes} onChange={(e) => setFiltroMes(e.target.value)}>
                  <option value="Todos">Todos los meses</option>
                  {mesesDisponibles.map((mes) => (
                    <option key={mes} value={mes}>
                      {new Date(`${mes}-01T12:00:00`).toLocaleDateString("es-PE", { month: "long", year: "numeric" })}
                    </option>
                  ))}
                </select>
                <select aria-label="Empresa" value={filtroEmpresa} onChange={(e) => setFiltroEmpresa(e.target.value)}><option>Todas</option>{empresas.map((item) => <option key={item}>{item}</option>)}</select>
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
              {soloJair && <div className="active-detail-filter"><span>Mostrando únicamente movimientos de Jair</span><button onClick={() => setSoloJair(false)}>Ver todos</button></div>}
              {lista.length ? (
                lista.map((movimiento) => (
                  <div className="movement" key={movimiento.id}>
                    <span className={`movement-icon ${movimiento.tipo}`}>
                      {movimiento.tipo === "ingreso" ? "↓" : movimiento.tipo === "transferencia" ? "⇄" : "↑"}
                    </span>
                    <div className="movement-name">
                      <b>{movimiento.concepto}</b>
                      <small>
                        {movimiento.tipo === "transferencia"
                          ? `${movimiento.empresaOrigen || movimiento.empresa} / ${movimiento.cuentaOrigen || movimiento.fuente} → ${movimiento.empresaDestino || movimiento.empresa} / ${movimiento.cuentaDestino || movimiento.detalleFuente || "Otro"} · ${movimiento.claseTransferencia || "Entre cuentas propias"}`
                          : `${movimiento.empresa} · ${movimiento.categoria}${movimiento.subcategoria ? ` / ${movimiento.subcategoria}` : ""} · ${movimiento.cuentaOrigen || movimiento.cuentaDestino || movimiento.fuente || "Otro"}`}{" "}
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
                      {movimiento.tipo === "ingreso" ? "+" : movimiento.tipo === "gasto" ? "−" : ""}
                      {soles(movimiento.monto)}
                    </strong>
                    {movimiento.tipo === "gasto" && movimiento.categoria === "Pasajes" && <button className={`budget-flag ${movimiento.excluirPresupuesto ? "active" : ""}`} onClick={() => alternarPresupuesto(movimiento.id)}>{movimiento.excluirPresupuesto ? "Por encargo" : "Excluir de meta"}</button>}
                    <button
                      className="edit-movement"
                      aria-label="Editar movimiento"
                      onClick={() => editar(movimiento)}
                    >
                      Editar
                    </button>
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
          </article>}
        </section>
      </section>

      {modalCuenta && (
        <div className="modal-backdrop" onMouseDown={() => setModalCuenta(false)}>
          <div className="modal account-modal" onMouseDown={(e) => e.stopPropagation()}>
            <button className="close" onClick={() => setModalCuenta(false)}>×</button>
            <span className="eyebrow">CUENTA Y RESPALDO</span>
            {session ? (
              <>
                <h2>Información sincronizada</h2>
                <p>Tus movimientos se guardan en la nube y estarán disponibles en tus otros dispositivos.</p>
                <div className="signed-account"><b>{session.user.email}</b><span>{estadoNube}</span></div>
                <button className="save logout" onClick={cerrarSesion}>Cerrar sesión</button>
              </>
            ) : (
              <>
                <h2>Guardar en la nube</h2>
                <p>Inicia sesión o crea una cuenta. Los movimientos de este dispositivo se migrarán automáticamente.</p>
                <form onSubmit={(e) => { e.preventDefault(); accederCuenta("ingresar"); }}>
                  <label>Correo electrónico<input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required /></label>
                  <label>Contraseña<input type="password" minLength={6} value={password} onChange={(e) => setPassword(e.target.value)} required /></label>
                  {mensajeCuenta && <p className="account-message">{mensajeCuenta}</p>}
                  <button className="save" disabled={procesandoCuenta}>Iniciar sesión</button>
                  <button type="button" className="create-account" disabled={procesandoCuenta} onClick={() => accederCuenta("crear")}>Crear una cuenta</button>
                </form>
              </>
            )}
          </div>
        </div>
      )}

      {modal && (
        <div className="modal-backdrop" onMouseDown={cerrarModalMovimiento}>
          <div className="modal" onMouseDown={(e) => e.stopPropagation()}>
            <button className="close" onClick={cerrarModalMovimiento}>
              ×
            </button>
            <span className="eyebrow">{movimientoEditando === null ? "NUEVO REGISTRO" : "EDITAR REGISTRO"}</span>
            <h2>{movimientoEditando === null ? "Registrar movimiento" : "Modificar movimiento"}</h2>
            <p>{movimientoEditando === null ? "Puedes cargar movimientos desde el 15 de junio de 2026." : "Los cambios actualizarán saldos, presupuestos y reportes."}</p>
            <form onSubmit={guardar}>
              <div className="type-toggle transfer-toggle">
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
                <button
                  type="button"
                  className={tipo === "transferencia" ? "selected transferencia" : ""}
                  onClick={() => setTipo("transferencia")}
                >
                  Transferencia
                </button>
              </div>
              {tipo !== "transferencia" && <label>
                Empresa o unidad
                <select value={empresa} onChange={(e) => setEmpresa(e.target.value as Empresa)}>
                  {empresas.map((item) => <option key={item}>{item}</option>)}
                </select>
              </label>}
              {tipo === "ingreso" && (
                <label>
                  Tipo de ingreso
                  <select value={categoriaIngreso} onChange={(e) => setCategoriaIngreso(e.target.value)}>
                    {categoriasIngreso.map((item) => <option key={item}>{item}</option>)}
                  </select>
                </label>
              )}
              {tipo === "gasto" && (
                <label>
                  Categoría
                  <select
                    autoFocus
                    value={categoria}
                    onChange={(e) => setCategoria(e.target.value)}
                  >
                    {categorias.map((item) => (
                      <option key={item}>{item}</option>
                    ))}
                  </select>
                </label>
              )}
              {tipo === "gasto" && categoria === "Extras" && (
                <label>
                  ¿En qué se fue este gasto?
                  <select value={subcategoria} onChange={(e) => setSubcategoria(e.target.value)}>
                    {subcategoriasExtras.map((item) => <option key={item}>{item}</option>)}
                  </select>
                </label>
              )}
              {tipo === "gasto" && categoria === "Pasajes" && (
                <label className="check-option">
                  <input type="checkbox" checked={excluirPresupuesto} onChange={(e) => setExcluirPresupuesto(e.target.checked)} />
                  <span><b>Pago por encargo o reembolsado</b><small>No se descontará de la meta mensual de pasajes.</small></span>
                </label>
              )}
              <label>
                Monto (S/)
                <input
                  autoFocus={tipo !== "gasto"}
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
              {tipo === "gasto" && (
              <label>
                  Concepto (opcional)
                  <input
                    placeholder="Ej. Almuerzo, pago de tarjeta"
                    value={concepto}
                    onChange={(e) => setConcepto(e.target.value)}
                  />
                </label>
              )}
              {tipo === "transferencia" ? (
                <>
                  <label>
                    Unidad de origen
                    <select value={empresaOrigen} onChange={(e) => { const valor = e.target.value as Empresa; setEmpresaOrigen(valor); setCuentaOrigen((cuentasPorEmpresa[valor] || [])[0] || ""); }}>
                      {empresas.map((item) => <option key={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Cuenta de origen
                    <select value={cuentaOrigen} onChange={(e) => setCuentaOrigen(e.target.value)}>
                      {(cuentasPorEmpresa[empresaOrigen] || []).map((item) => <option key={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Unidad de destino
                    <select value={empresaDestino} onChange={(e) => { const valor = e.target.value as Empresa; setEmpresaDestino(valor); setCuentaDestino((cuentasPorEmpresa[valor] || [])[0] || ""); }}>
                      {empresas.map((item) => <option key={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Cuenta de destino
                    <select value={cuentaDestino} onChange={(e) => setCuentaDestino(e.target.value)}>
                      {(cuentasPorEmpresa[empresaDestino] || []).map((item) => <option key={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Motivo de la transferencia
                    <select value={claseTransferencia} onChange={(e) => setClaseTransferencia(e.target.value)}>
                      {clasesTransferencia.map((item) => <option key={item}>{item}</option>)}
                    </select>
                  </label>
                </>
              ) : (
                <>
                  <label>
                    Medio de pago o fuente
                    <select value={fuente} onChange={(e) => { setFuente(e.target.value); setDetalleFuente(""); }}>
                      {fuentes.map((item) => <option key={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Detalle de la fuente
                    <input
                      placeholder={fuente === "Efectivo" || fuente === "Otro" ? "Escribe el detalle" : "Se completa automáticamente"}
                      value={detalleAutomatico(fuente) || detalleFuente}
                      readOnly={Boolean(detalleAutomatico(fuente))}
                      onChange={(e) => setDetalleFuente(e.target.value)}
                    />
                  </label>
                </>
              )}
              {tipo === "ingreso" && (
                <label>
                  Concepto (opcional)
                  <input
                    placeholder="Ej. Sueldo, devolución, depósito"
                    value={concepto}
                    onChange={(e) => setConcepto(e.target.value)}
                  />
                </label>
              )}
              {tipo === "transferencia" && (
                <label>
                  Concepto (opcional)
                  <input
                    placeholder="Ej. Traspaso a cuenta de ahorro"
                    value={concepto}
                    onChange={(e) => setConcepto(e.target.value)}
                  />
                </label>
              )}
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
                {movimientoEditando === null ? "Guardar movimiento" : "Guardar cambios"}
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
              Los archivos se procesan en este dispositivo. Puedes seleccionar
              varios a la vez y confirmar el mapeo de cada uno.
            </p>

            <label
              className={`file-drop ${archivosExcel.length ? "loaded" : ""}`}
            >
              <input
                ref={inputArchivo}
                type="file"
                accept=".xlsx,.csv"
                multiple
                onChange={cargarExcel}
              />
              <span>{archivosExcel.length ? "＋" : "⇧"}</span>
              <b>
                {archivosExcel.length
                  ? "Agregar más archivos"
                  : "Seleccionar uno o varios archivos"}
              </b>
              <small>
                {archivosExcel.length
                  ? `${archivosExcel.length} ${
                      archivosExcel.length === 1 ? "archivo" : "archivos"
                    } · ${totalFilasExcel} filas encontradas`
                  : "Formatos permitidos: .xlsx y .csv"}
              </small>
            </label>

            {archivosExcel.length > 0 && archivoActivo && (
              <>
                <div className="file-tabs">
                  {archivosExcel.map((archivo, indice) => (
                    <div
                      className={indice === indiceArchivo ? "active" : ""}
                      key={archivo.id}
                    >
                      <button
                        type="button"
                        onClick={() => setIndiceArchivo(indice)}
                      >
                        <b>{archivo.nombre}</b>
                        <small>{archivo.filas.length} filas</small>
                      </button>
                      <button
                        type="button"
                        aria-label={`Quitar ${archivo.nombre}`}
                        onClick={() => quitarArchivo(archivo.id)}
                      >
                        ×
                      </button>
                    </div>
                  ))}
                </div>

                <div className="mapping-title">
                  <div>
                    <b>Mapeo de columnas</b>
                    <span>{archivoActivo.nombre}</span>
                  </div>
                  <small>
                    Archivo {indiceArchivo + 1} de {archivosExcel.length}
                  </small>
                </div>
                <div className="mapping-grid">
                  <label>
                    Fecha *
                    <select
                      value={archivoActivo.mapa.fecha}
                      onChange={(e) =>
                        actualizarMapaActivo("fecha", e.target.value)
                      }
                    >
                      <option value="">Seleccionar columna</option>
                      {archivoActivo.columnas.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Descripción *
                    <select
                      value={archivoActivo.mapa.concepto}
                      onChange={(e) =>
                        actualizarMapaActivo("concepto", e.target.value)
                      }
                    >
                      <option value="">Seleccionar columna</option>
                      {archivoActivo.columnas.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Monto único
                    <select
                      value={archivoActivo.mapa.monto}
                      onChange={(e) =>
                        actualizarMapaActivo("monto", e.target.value)
                      }
                    >
                      <option value="">No aplica</option>
                      {archivoActivo.columnas.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Cargo / débito
                    <select
                      value={archivoActivo.mapa.cargo}
                      onChange={(e) =>
                        actualizarMapaActivo("cargo", e.target.value)
                      }
                    >
                      <option value="">No aplica</option>
                      {archivoActivo.columnas.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Abono / crédito
                    <select
                      value={archivoActivo.mapa.abono}
                      onChange={(e) =>
                        actualizarMapaActivo("abono", e.target.value)
                      }
                    >
                      <option value="">No aplica</option>
                      {archivoActivo.columnas.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Tipo
                    <select
                      value={archivoActivo.mapa.tipo}
                      onChange={(e) =>
                        actualizarMapaActivo("tipo", e.target.value)
                      }
                    >
                      <option value="">No aplica</option>
                      {archivoActivo.columnas.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Medio o cuenta
                    <select
                      value={archivoActivo.mapa.fuente}
                      onChange={(e) =>
                        actualizarMapaActivo("fuente", e.target.value)
                      }
                    >
                      <option value="">No aplica</option>
                      {archivoActivo.columnas.map((columna) => (
                        <option key={columna}>{columna}</option>
                      ))}
                    </select>
                  </label>
                  {!archivoActivo.mapa.tipo &&
                    !archivoActivo.mapa.cargo &&
                    !archivoActivo.mapa.abono && (
                      <label>
                        Los montos positivos son
                        <select
                          value={archivoActivo.positivosSon}
                          onChange={(e) =>
                            actualizarPositivosActivo(e.target.value as Tipo)
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
                  {archivoActivo.filas.slice(0, 3).map((fila, indice) => (
                    <div key={indice}>
                      <span>
                        {String(
                          valorFila(fila, archivoActivo.mapa.fecha) ||
                            "Sin fecha",
                        )}
                      </span>
                      <strong>
                        {String(
                          valorFila(fila, archivoActivo.mapa.concepto) ||
                            "Sin descripción",
                        )}
                      </strong>
                      <span>
                        {String(
                          valorFila(
                            fila,
                            archivoActivo.mapa.monto ||
                              archivoActivo.mapa.cargo ||
                              archivoActivo.mapa.abono,
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
              disabled={!archivosExcel.length}
              onClick={prepararRevision}
            >
              Revisar {totalFilasExcel} movimientos
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
              <label>
                Empresa o unidad
                <select
                  value={actual.empresa}
                  onChange={(e) =>
                    actualizarActual({
                      empresa: e.target.value as Empresa,
                    })
                  }
                >
                  {empresas.map((item) => (
                    <option key={item}>{item}</option>
                  ))}
                </select>
              </label>
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
                <label>
                  Tipo de ingreso
                  <select
                    value={actual.categoria}
                    onChange={(e) =>
                      actualizarActual({ categoria: e.target.value })
                    }
                  >
                    {categoriasIngreso.map((item) => (
                      <option key={item}>{item}</option>
                    ))}
                  </select>
                </label>
              )}
              {actual.tipo === "gasto" && actual.categoria === "Extras" && (
                <label>
                  Detalle de Extras
                  <select
                    value={actual.subcategoria || "Otros"}
                    onChange={(e) =>
                      actualizarActual({ subcategoria: e.target.value })
                    }
                  >
                    {subcategoriasExtras.map((item) => (
                      <option key={item}>{item}</option>
                    ))}
                  </select>
                </label>
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
