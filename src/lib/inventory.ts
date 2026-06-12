export type TipoMovimiento = "entrada" | "salida";
export type TipoReporte = "todos" | "entradas" | "salidas";
export type EstadoStock = "AGOTADO" | "CRÍTICO" | "BAJO" | "ÓPTIMO";
export type Vista = "dashboard" | "inventario" | "movimientos";

export type Producto = {
  id: number;
  nombre: string;
  categoria: Categoria;
  stock: number;
  stock_minimo: number;
  cantidad_base: number;
  ubicacion?: string;
  created_at?: string;
};

export type Movimiento = {
  id: number;
  producto_id: number;
  tipo: TipoMovimiento;
  cantidad: number;
  motivo: string;
  proveedor?: string;
  fecha?: string;
  factura?: string;
  notas?: string;
  created_at: string;
};

export type Usuario = {
  id: number;
  nombre: string;
  rol?: "admin" | "usuario";
};

export type Notificacion = {
  tipo: "success" | "error" | "warning" | "info";
  mensaje: string;
};

export enum Categoria {
  Cervezas = "Cervezas",
  SodaGinger = "Soda & Ginger",
  SodaHatsu = "Soda Hatsu",
  Gaseosas = "Gaseosas",
  Vinos = "Vinos",
  Agua = "Agua",
}

export const CATEGORIAS = Object.values(Categoria);

export const CATEGORIA_COLOR: Record<Categoria, string> = {
  [Categoria.Cervezas]: "#fbbf24",
  [Categoria.SodaGinger]: "#60a5fa",
  [Categoria.SodaHatsu]: "#c084fc",
  [Categoria.Gaseosas]: "#4ade80",
  [Categoria.Vinos]: "#f472b6",
  [Categoria.Agua]: "#22d3ee",
};

export const CATEGORIA_BG: Record<Categoria, string> = {
  [Categoria.Cervezas]: "#1a1200",
  [Categoria.SodaGinger]: "#0a1a2a",
  [Categoria.SodaHatsu]: "#1a0a1a",
  [Categoria.Gaseosas]: "#0a1a0a",
  [Categoria.Vinos]: "#1a0a10",
  [Categoria.Agua]: "#0a1a1f",
};

export const CATEGORIA_BORDER: Record<Categoria, string> = {
  [Categoria.Cervezas]: "#92400e",
  [Categoria.SodaGinger]: "#1e40af",
  [Categoria.SodaHatsu]: "#6d28d9",
  [Categoria.Gaseosas]: "#166534",
  [Categoria.Vinos]: "#9d174d",
  [Categoria.Agua]: "#0e7490",
};

export const CATEGORIA_EMOJI: Record<Categoria, string> = {
  [Categoria.Cervezas]: "🍺",
  [Categoria.SodaGinger]: "🫚",
  [Categoria.SodaHatsu]: "💜",
  [Categoria.Gaseosas]: "🥤",
  [Categoria.Vinos]: "🍷",
  [Categoria.Agua]: "💧",
};

export function estadoInfo(p: Producto): {
  label: EstadoStock;
  color: string;
  bg: string;
  dot: string;
} {
  if (p.stock === 0)
    return { label: "AGOTADO", color: "#ef4444", bg: "#2a0a0a", dot: "#ef4444" };
  if (p.stock <= p.stock_minimo)
    return { label: "CRÍTICO", color: "#f97316", bg: "#2a1a0a", dot: "#f97316" };
  if (p.stock <= p.stock_minimo * 1.5)
    return { label: "BAJO", color: "#eab308", bg: "#1a1200", dot: "#eab308" };
  return { label: "ÓPTIMO", color: "#22c55e", bg: "#0a2a14", dot: "#22c55e" };
}

export function formatFecha(fecha: string): string {
  return new Date(fecha).toLocaleDateString("es-CO", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

export function validateCantidad(cantidad: string): { valida: boolean; error?: string } {
  const num = parseInt(cantidad, 10);
  if (!cantidad) return { valida: false, error: "La cantidad es requerida" };
  if (isNaN(num)) return { valida: false, error: "Ingrese un número válido" };
  if (num <= 0) return { valida: false, error: "La cantidad debe ser mayor a 0" };
  return { valida: true };
}

export function calcularPedidoSugerido(stock: number, cantidadBase: number): number {
  const deficit = cantidadBase - stock;
  return deficit > 0 ? deficit : 0;
}

/** Corrige el bug: TipoReporte usa "entradas"/"salidas" pero TipoMovimiento es "entrada"/"salida" */
export function coincideTipoFiltro(tipoMovimiento: TipoMovimiento, filtro: TipoReporte): boolean {
  if (filtro === "todos") return true;
  if (filtro === "entradas") return tipoMovimiento === "entrada";
  if (filtro === "salidas") return tipoMovimiento === "salida";
  return true;
}

export const CHART_TOOLTIP_STYLE = {
  background: "#0d1117",
  border: "1px solid #243044",
  borderRadius: "8px",
  color: "#f1f5f9",
};
