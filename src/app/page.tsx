"use client";

import { supabase } from "@/lib/supabase";
import { useEffect, useState, useMemo, useCallback, createContext, useContext } from "react";
import * as ExcelJS from "exceljs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, AreaChart, Area
} from 'recharts';

// ==================== TIPOS Y ENUMS ====================

type TipoMovimiento = 'entrada' | 'salida';
type TipoReporte = 'todos' | 'entradas' | 'salidas';
type EstadoStock = 'AGOTADO' | 'CRÍTICO' | 'BAJO' | 'ÓPTIMO';
type Vista = "dashboard" | "inventario" | "movimientos";
type Tema = "dark" | "light";

type Producto = {
  id: number;
  nombre: string;
  categoria: Categoria;
  stock: number;
  stock_minimo: number;
  cantidad_base: number;
  unidades_por_caja: number;  // NUEVO: cuántas unidades vienen por caja/paca
  ubicacion?: string;
  created_at?: string;
  ultima_actualizacion?: string;
};

type Movimiento = {
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
  usuario?: string;
};

type Usuario = {
  id: number;
  nombre: string;
  rol?: 'admin' | 'usuario';
};

type Notificacion = {
  tipo: 'success' | 'error' | 'warning' | 'info';
  mensaje: string;
  id?: number;
};

type AccionHistorial = {
  id: number;
  usuario: string;
  accion: string;
  detalles: string;
  fecha: string;
};

enum Categoria {
  Cervezas = "Cervezas",
  SodaGinger = "Soda & Ginger",
  SodaHatsu = "Soda Hatsu",
  Gaseosas = "Gaseosas",
  Vinos = "Vinos",
  Agua = "Agua"
}

const CATEGORIAS = Object.values(Categoria);

const CATEGORIA_COLOR: Record<Categoria, string> = {
  [Categoria.Cervezas]: "#fbbf24",
  [Categoria.SodaGinger]: "#60a5fa",
  [Categoria.SodaHatsu]: "#c084fc",
  [Categoria.Gaseosas]: "#4ade80",
  [Categoria.Vinos]: "#f472b6",
  [Categoria.Agua]: "#60a5fa",
};

const CATEGORIA_BG: Record<Categoria, string> = {
  [Categoria.Cervezas]: "#1a1200",
  [Categoria.SodaGinger]: "#0a1a2a",
  [Categoria.SodaHatsu]: "#1a0a1a",
  [Categoria.Gaseosas]: "#0a1a0a",
  [Categoria.Vinos]: "#1a0a10",
  [Categoria.Agua]: "#0a1a2a",
};

const CATEGORIA_BORDER: Record<Categoria, string> = {
  [Categoria.Cervezas]: "#92400e",
  [Categoria.SodaGinger]: "#1e40af",
  [Categoria.SodaHatsu]: "#6d28d9",
  [Categoria.Gaseosas]: "#166534",
  [Categoria.Vinos]: "#9d174d",
  [Categoria.Agua]: "#1e40af",
};

const CATEGORIA_EMOJI: Record<Categoria, string> = {
  [Categoria.Cervezas]: "🍺",
  [Categoria.SodaGinger]: "🫚",
  [Categoria.SodaHatsu]: "💜",
  [Categoria.Gaseosas]: "🥤",
  [Categoria.Vinos]: "🍷",
  [Categoria.Agua]: "💧",
};

// ==================== CONTEXTO DE TEMA ====================

const TemaContext = createContext<{ tema: Tema; toggleTema: () => void }>({ tema: "dark", toggleTema: () => {} });

export function useTema() {
  return useContext(TemaContext);
}

// ==================== COMPONENTE DE NOTIFICACIONES ====================

function NotificationToast({ notificacion, onClose }: { notificacion: Notificacion | null; onClose: () => void }) {
  const [mounted, setMounted] = useState(false);

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    if (notificacion) {
      const timer = setTimeout(onClose, 3000);
      return () => clearTimeout(timer);
    }
  }, [notificacion, onClose]);

  if (!mounted || !notificacion) return null;

  const colors = {
    success: "bg-green-500/20 border-green-500/30 text-green-400",
    error: "bg-red-500/20 border-red-500/30 text-red-400",
    warning: "bg-yellow-500/20 border-yellow-500/30 text-yellow-400",
    info: "bg-blue-500/20 border-blue-500/30 text-blue-400"
  };

  const icons = {
    success: "✅",
    error: "❌",
    warning: "⚠️",
    info: "ℹ️"
  };

  return (
    <div className="fixed top-20 right-6 z-50 animate-in slide-in-from-right-5 duration-300">
      <div className={`px-6 py-4 rounded-xl border backdrop-blur-md ${colors[notificacion.tipo]} shadow-2xl flex items-center gap-3 min-w-[300px]`}>
        <span className="text-2xl">{icons[notificacion.tipo]}</span>
        <p className="text-sm font-medium">{notificacion.mensaje}</p>
        <button onClick={onClose} className="ml-auto text-slate-400 hover:text-white transition-colors">✕</button>
      </div>
    </div>
  );
}

// ==================== FUNCIONES UTILITARIAS ====================

function estadoInfo(p: Producto): { label: EstadoStock; color: string; bg: string; dot: string; textColor: string } {
  if (p.stock === 0) return { label: "AGOTADO", color: "#ef4444", bg: "#2a0a0a", dot: "#ef4444", textColor: "#fca5a5" };
  if (p.stock <= p.stock_minimo) return { label: "CRÍTICO", color: "#f97316", bg: "#2a1a0a", dot: "#f97316", textColor: "#fdba74" };
  if (p.stock <= p.stock_minimo * 1.5) return { label: "BAJO", color: "#eab308", bg: "#1a1200", dot: "#eab308", textColor: "#fde047" };
  return { label: "ÓPTIMO", color: "#22c55e", bg: "#0a2a14", dot: "#22c55e", textColor: "#86efac" };
}

function formatFecha(fecha: string): string {
  return new Date(fecha).toLocaleDateString("es-CO", { 
    day: "numeric", 
    month: "short", 
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatFechaCorta(fecha: string): string {
  return new Date(fecha).toLocaleDateString("es-CO", { 
    day: "numeric", 
    month: "short"
  });
}

function validateCantidad(cantidad: string): { valida: boolean; error?: string } {
  const num = parseInt(cantidad);
  if (!cantidad) return { valida: false, error: "La cantidad es requerida" };
  if (isNaN(num)) return { valida: false, error: "Ingrese un número válido" };
  if (num <= 0) return { valida: false, error: "La cantidad debe ser mayor a 0" };
  return { valida: true };
}

function calcularPedidoSugerido(stock: number, cantidadBase: number): number {
  const deficit = cantidadBase - stock;
  return deficit > 0 ? deficit : 0;
}

// NUEVA FUNCIÓN: Calcular cajas necesarias
function calcularCajasNecesarias(unidades: number, unidadesPorCaja: number): { cajas: number; resto: number } {
  if (!unidadesPorCaja || unidadesPorCaja <= 0) return { cajas: 0, resto: unidades };
  const cajas = Math.floor(unidades / unidadesPorCaja);
  const resto = unidades % unidadesPorCaja;
  return { cajas, resto };
}

function calcularDiasCobertura(stock: number, consumoPromedioDiario: number): number {
  if (consumoPromedioDiario <= 0) return 999;
  return Math.floor(stock / consumoPromedioDiario);
}

// ==================== LOGIN ====================

function LoginScreen({ onLogin }: { onLogin: (usuario: Usuario) => void }) {
  const [nombre, setNombre] = useState("");
  const [contrasena, setContrasena] = useState("");
  const [error, setError] = useState("");
  const [cargando, setCargando] = useState(false);
  
  const handleLogin = async (e: React.FormEvent) => { 
    e.preventDefault(); 
    if (!nombre.trim() || !contrasena.trim()) { 
      setError("Completa todos los campos"); 
      return; 
    } 
    setCargando(true); 
    setError(""); 
    
    if (nombre.trim() === "Bartender" && contrasena === "123456") {
      const usuarioBartender: Usuario = { id: 1, nombre: "Bartender", rol: "admin" };
      localStorage.setItem("usuario_activo", JSON.stringify(usuarioBartender)); 
      onLogin(usuarioBartender);
      setCargando(false);
      return;
    }
    
    setError("Usuario o contraseña incorrectos"); 
    setCargando(false);
  };
  
  return (
    <div className="min-h-screen bg-gradient-to-br from-[#0a0d16] to-[#141824] flex items-center justify-center p-6">
      <div className="w-full max-w-md animate-in zoom-in-95 duration-500">
        <div className="text-center mb-10">
          <div className="inline-flex items-center justify-center w-24 h-24 rounded-3xl bg-gradient-to-br from-amber-500 to-amber-600 mb-6 shadow-2xl animate-bounce-slow">
            <span className="text-5xl">🍺</span>
          </div>
          <h1 className="text-4xl font-bold bg-gradient-to-r from-amber-400 to-amber-600 bg-clip-text text-transparent">INVENTARIO</h1>
          <p className="text-sm text-slate-500 mt-3">Sistema profesional de control de inventario</p>
        </div>
        
        <div className="bg-[#0f1117]/80 backdrop-blur-sm border border-[#2a2a3e] rounded-2xl shadow-2xl overflow-hidden">
          <div className="px-8 pt-8 pb-5 border-b border-[#2a2a3e]">
            <h2 className="text-xl font-semibold text-white">Iniciar sesión</h2>
            <p className="text-sm text-slate-500 mt-1">Ingresa tus credenciales para acceder al sistema</p>
          </div>
          
          <form onSubmit={handleLogin} className="px-8 pb-8 pt-6 space-y-6">
            {error && (
              <div className="bg-red-500/10 border border-red-500/30 rounded-xl p-4 animate-shake">
                <p className="text-sm text-red-400">{error}</p>
              </div>
            )}
            
            <div>
              <label className="block text-sm font-medium text-slate-400 mb-2">Usuario</label>
              <input 
                type="text" 
                value={nombre} 
                onChange={(e) => setNombre(e.target.value)} 
                placeholder="Nombre de usuario" 
                className="w-full px-4 py-3 bg-[#0a0d16] border border-[#2a2a3e] rounded-xl text-white text-base placeholder-slate-600 focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20 transition-all"
                autoFocus 
              />
            </div>
            
            <div>
              <label className="block text-sm font-medium text-slate-400 mb-2">Contraseña</label>
              <input 
                type="password" 
                value={contrasena} 
                onChange={(e) => setContrasena(e.target.value)} 
                placeholder="••••••••" 
                className="w-full px-4 py-3 bg-[#0a0d16] border border-[#2a2a3e] rounded-xl text-white text-base placeholder-slate-600 focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20 transition-all"
              />
            </div>
            
            <button 
              type="submit" 
              disabled={cargando} 
              className="w-full py-3 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white font-semibold rounded-xl transition-all duration-200 disabled:opacity-50 flex items-center justify-center gap-2 text-base shadow-lg"
            >
              {cargando ? (
                <>
                  <div className="animate-spin rounded-full h-5 w-5 border-2 border-white border-t-transparent"></div>
                  Ingresando...
                </>
              ) : "Ingresar al sistema"}
            </button>
          </form>
          
          <div className="bg-[#0a0d16] px-8 py-4 border-t border-[#2a2a3e]">
            <div className="flex items-center justify-between text-sm text-slate-500">
              <div className="flex items-center gap-2">
                <span className="w-2 h-2 bg-green-500 rounded-full animate-pulse"></span>
                <span>Demo activo</span>
              </div>
              <div className="flex gap-4">
                <span>Usuario: <span className="text-amber-400 font-mono">Bartender</span></span>
                <span>Contraseña: <span className="text-amber-400 font-mono">123456</span></span>
              </div>
            </div>
          </div>
        </div>
        
        <p className="text-center text-xs text-slate-600 mt-8">
          © {new Date().getFullYear()} INVENTARIO - Sistema de Gestión de Inventario
        </p>
      </div>
    </div>
  );
}

// ==================== MODALES ====================

type ModalData = {
  tipo: "total" | "critico" | "agotados" | "optimo" | "pedidos" | null;
};

function ModalDetalle({ tipo, productos, onClose }: { tipo: string; productos: Producto[]; onClose: () => void }) {
  if (!tipo) return null;

  let titulo = "";
  let descripcion = "";
  let icono = "";
  let productosFiltrados: Producto[] = [];
  let color = "";

  switch (tipo) {
    case "total":
      titulo = "Todos los Productos";
      descripcion = "Listado completo del inventario";
      icono = "📦";
      productosFiltrados = productos;
      color = "#3b82f6";
      break;
    case "critico":
      titulo = "Stock Crítico";
      descripcion = "Productos que requieren reposición inmediata";
      icono = "⚠️";
      productosFiltrados = productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0);
      color = "#f97316";
      break;
    case "agotados":
      titulo = "Productos Agotados";
      descripcion = "Productos sin stock disponible";
      icono = "🚫";
      productosFiltrados = productos.filter(p => p.stock === 0);
      color = "#ef4444";
      break;
    case "optimo":
      titulo = "Stock Óptimo";
      descripcion = "Productos con nivel de stock adecuado";
      icono = "✅";
      productosFiltrados = productos.filter(p => p.stock > p.stock_minimo * 1.5);
      color = "#22c55e";
      break;
    case "pedidos":
      titulo = "Pedido Sugerido";
      descripcion = "Productos que necesitan reposición para alcanzar la cantidad base";
      icono = "📦";
      productosFiltrados = productos.filter(p => calcularPedidoSugerido(p.stock, p.cantidad_base) > 0);
      color = "#f97316";
      break;
    default:
      return null;
  }

  const totalUnidades = productosFiltrados.reduce((sum, p) => sum + p.stock, 0);
  const totalPedido = productosFiltrados.reduce((sum, p) => sum + calcularPedidoSugerido(p.stock, p.cantidad_base), 0);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6 animate-in fade-in duration-200" onClick={onClose}>
      <div className="relative w-full max-w-5xl max-h-[85vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] shadow-2xl animate-in zoom-in-95 duration-300" onClick={(e) => e.stopPropagation()}>
        <div className="sticky top-0 bg-[#0f1117] border-b border-[#2a2a3e] p-6 flex justify-between items-center z-10">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center text-2xl" style={{ background: `${color}20`, color }}>{icono}</div>
            <div>
              <h2 className="text-xl font-bold text-white">{titulo}</h2>
              <p className="text-sm text-slate-500">{descripcion}</p>
            </div>
          </div>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>

        <div className="grid grid-cols-2 gap-6 p-6 bg-[#0a0d16] border-b border-[#2a2a3e]">
          <div className="text-center p-4 rounded-xl bg-[#0f1117]">
            <div className="text-4xl font-bold mb-1" style={{ color }}>{productosFiltrados.length}</div>
            <div className="text-sm text-slate-500 uppercase tracking-wide">Productos</div>
          </div>
          <div className="text-center p-4 rounded-xl bg-[#0f1117]">
            <div className="text-4xl font-bold mb-1" style={{ color }}>{tipo === "pedidos" ? totalPedido : totalUnidades}</div>
            <div className="text-sm text-slate-500 uppercase tracking-wide">{tipo === "pedidos" ? "Unidades a Pedir" : "Unidades Totales"}</div>
          </div>
        </div>

        <div className="p-6">
          <div className="overflow-x-auto">
            <table className="w-full">
              <thead className="border-b border-[#2a2a3e]">
                <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wide">
                  <th className="pb-4">Producto</th>
                  <th className="pb-4">Categoría</th>
                  <th className="pb-4 text-right">Stock Actual</th>
                  {tipo === "pedidos" && (
                    <>
                      <th className="pb-4 text-right">Cantidad Base</th>
                      <th className="pb-4 text-right">Unid./Caja</th>
                      <th className="pb-4 text-right">Pedido (Cajas)</th>
                    </>
                  )}
                  {tipo !== "pedidos" && (
                    <th className="pb-4 text-right">Stock Mínimo</th>
                  )}
                  <th className="pb-4 text-right">Ubicación</th>
                </tr>
              </thead>
              <tbody>
                {productosFiltrados.map((p, idx) => {
                  const pedido = calcularPedidoSugerido(p.stock, p.cantidad_base);
                  const { cajas, resto } = calcularCajasNecesarias(pedido, p.unidades_por_caja || 1);
                  const textoPedido = cajas > 0 ? `${cajas} caja${cajas !== 1 ? 's' : ''}${resto > 0 ? ` + ${resto} und` : ''}` : `${resto} und`;
                  return (
                    <tr key={p.id} className="border-b border-[#1a1a2a] hover:bg-[#1a1a2a] transition-colors">
                      <td className="py-3 text-white font-medium">{p.nombre}</td>
                      <td className="py-3"><span className="text-xs px-3 py-1.5 rounded-full" style={{ background: CATEGORIA_BG[p.categoria], color: CATEGORIA_COLOR[p.categoria], border: `1px solid ${CATEGORIA_BORDER[p.categoria]}` }}>{CATEGORIA_EMOJI[p.categoria]} {p.categoria}</span></td>
                      <td className="py-3 text-right text-xl font-bold" style={{ color }}>{p.stock}</td>
                      {tipo === "pedidos" && (
                        <>
                          <td className="py-3 text-right text-lg font-semibold text-blue-400">{p.cantidad_base}</td>
                          <td className="py-3 text-right text-lg font-semibold text-purple-400">{p.unidades_por_caja || 1}</td>
                          <td className="py-3 text-right text-lg font-bold" style={{ color: "#f97316" }}>{textoPedido}</td>
                        </>
                      )}
                      {tipo !== "pedidos" && (
                        <td className="py-3 text-right text-slate-400">{p.stock_minimo}</td>
                      )}
                      <td className="py-3 text-right text-slate-500">{p.ubicacion || "—"}</td>
                    </tr>
                  );
                })}
                {productosFiltrados.length === 0 && (
                  <tr>
                    <td colSpan={tipo === "pedidos" ? 7 : 5} className="text-center py-12 text-slate-500">
                      {tipo === "pedidos" ? "No hay productos que necesiten pedido" : "No hay productos en esta categoría"}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

        {tipo === "pedidos" && productosFiltrados.length > 0 && (
          <div className="p-6 border-t border-[#2a2a3e] bg-[#0a0d16]">
            <div className="flex justify-between items-center flex-wrap gap-4">
              <div>
                <span className="text-sm text-slate-400">Total de unidades a pedir: </span>
                <span className="text-2xl font-bold text-orange-400">{totalPedido}</span>
              </div>
              <button 
                onClick={() => {
                  const pedidoText = productosFiltrados.map(p => {
                    const pedido = calcularPedidoSugerido(p.stock, p.cantidad_base);
                    const { cajas, resto } = calcularCajasNecesarias(pedido, p.unidades_por_caja || 1);
                    const texto = cajas > 0 ? `${cajas} caja${cajas !== 1 ? 's' : ''}${resto > 0 ? ` + ${resto} und` : ''}` : `${resto} und`;
                    return `${p.nombre}: ${texto}`;
                  }).join("\n");
                  alert(`Pedido sugerido generado:\nTotal: ${totalPedido} unidades\n\n${pedidoText}`);
                }}
                className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-orange-600 to-orange-500 hover:from-orange-500 hover:to-orange-600 text-white font-medium transition-all flex items-center gap-2 shadow-lg"
              >
                📋 Generar Orden de Pedido
              </button>
            </div>
          </div>
        )}

        <div className="p-6 border-t border-[#2a2a3e] text-right">
          <button onClick={onClose} className="px-6 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium transition-all">Cerrar</button>
        </div>
      </div>
    </div>
  );
}

function ModalEstadisticas({ productos, movimientos, onClose }: { productos: Producto[]; movimientos: Movimiento[]; onClose: () => void }) {
  const consumoPorCategoria = CATEGORIAS.map(cat => ({
    name: cat,
    stock: productos.filter(p => p.categoria === cat).reduce((sum, p) => sum + p.stock, 0),
  })).filter(item => item.stock > 0);

  const ultimos7Dias = useMemo(() => {
    return Array.from({ length: 7 }, (_, i) => {
      const fecha = new Date();
      fecha.setDate(fecha.getDate() - i);
      fecha.setHours(0, 0, 0, 0);
      const fechaStr = fecha.toISOString().split('T')[0];
      const movimientosDia = movimientos.filter(m => m.created_at?.split('T')[0] === fechaStr);
      return { 
        fecha: formatFechaCorta(fechaStr), 
        entradas: movimientosDia.filter(m => m.tipo === "entrada").reduce((sum, m) => sum + m.cantidad, 0), 
        salidas: movimientosDia.filter(m => m.tipo === "salida").reduce((sum, m) => sum + m.cantidad, 0) 
      };
    }).reverse();
  }, [movimientos]);

  const ultimos30Dias = useMemo(() => {
    return Array.from({ length: 30 }, (_, i) => {
      const fecha = new Date();
      fecha.setDate(fecha.getDate() - i);
      fecha.setHours(0, 0, 0, 0);
      const fechaStr = fecha.toISOString().split('T')[0];
      const movimientosDia = movimientos.filter(m => m.created_at?.split('T')[0] === fechaStr);
      return { 
        fecha: formatFechaCorta(fechaStr), 
        entradas: movimientosDia.filter(m => m.tipo === "entrada").reduce((sum, m) => sum + m.cantidad, 0), 
        salidas: movimientosDia.filter(m => m.tipo === "salida").reduce((sum, m) => sum + m.cantidad, 0) 
      };
    }).reverse();
  }, [movimientos]);

  const topProductos = useMemo(() => {
    return productos.map(p => {
      const salidas = movimientos.filter(m => m.producto_id === p.id && m.tipo === "salida").reduce((sum, m) => sum + m.cantidad, 0);
      return { nombre: p.nombre, salidas, stock: p.stock, categoria: p.categoria };
    }).sort((a, b) => b.salidas - a.salidas).slice(0, 10);
  }, [productos, movimientos]);

  const COLORS = ['#fbbf24', '#60a5fa', '#c084fc', '#4ade80', '#f472b6', '#f97316', '#a855f7', '#ec4899', '#14b8a6', '#f43f5e'];
  
  const rotacionInventario = useMemo(() => {
    const salidasUltimoMes = movimientos
      .filter(m => m.tipo === 'salida' && 
        new Date(m.created_at) > new Date(Date.now() - 30 * 24 * 60 * 60 * 1000))
      .reduce((sum, m) => sum + m.cantidad, 0);
    const promedioInventario = productos.reduce((sum, p) => sum + p.stock, 0) / productos.length;
    return promedioInventario > 0 ? (salidasUltimoMes / promedioInventario).toFixed(2) : "0";
  }, [movimientos, productos]);

  const consumoPromedioDiario = useMemo(() => {
    const salidasUltimoMes = movimientos
      .filter(m => m.tipo === 'salida' && 
        new Date(m.created_at) > new Date(Date.now() - 30 * 24 * 60 * 60 * 1000))
      .reduce((sum, m) => sum + m.cantidad, 0);
    return salidasUltimoMes / 30;
  }, [movimientos]);

  const valorInventario = useMemo(() => {
    // Esto es un estimado - idealmente tendrías precios en la DB
    return productos.reduce((sum, p) => sum + (p.stock * 15000), 0);
  }, [productos]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6 animate-in fade-in duration-200" onClick={onClose}>
      <div className="relative w-full max-w-6xl max-h-[90vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl animate-in zoom-in-95 duration-300" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-6 sticky top-0 bg-[#0f1117] z-10 pb-4 border-b border-[#2a2a3e]">
          <h2 className="text-2xl font-bold bg-gradient-to-r from-amber-400 to-amber-600 bg-clip-text text-transparent">📊 Estadísticas y Análisis</h2>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>
        
        {/* Métricas clave */}
        <div className="grid grid-cols-4 gap-4 mb-8">
          <div className="bg-gradient-to-br from-blue-500/10 to-blue-600/5 rounded-xl p-4 text-center border border-blue-500/30">
            <div className="text-2xl font-bold text-blue-400">{rotacionInventario}x</div>
            <div className="text-xs text-slate-500">Rotación (30 días)</div>
          </div>
          <div className="bg-gradient-to-br from-green-500/10 to-green-600/5 rounded-xl p-4 text-center border border-green-500/30">
            <div className="text-2xl font-bold text-green-400">{consumoPromedioDiario.toFixed(1)}</div>
            <div className="text-xs text-slate-500">Consumo diario promedio</div>
          </div>
          <div className="bg-gradient-to-br from-purple-500/10 to-purple-600/5 rounded-xl p-4 text-center border border-purple-500/30">
            <div className="text-2xl font-bold text-purple-400">{movimientos.length}</div>
            <div className="text-xs text-slate-500">Total Movimientos</div>
          </div>
          <div className="bg-gradient-to-br from-amber-500/10 to-amber-600/5 rounded-xl p-4 text-center border border-amber-500/30">
            <div className="text-2xl font-bold text-amber-400">${(valorInventario / 1000000).toFixed(1)}M</div>
            <div className="text-xs text-slate-500">Valor estimado inventario</div>
          </div>
        </div>
        
        {/* Gráfico de stock por categoría */}
        <div className="mb-8">
          <h3 className="text-base font-semibold text-slate-300 mb-4 flex items-center gap-2">📊 Stock por Categoría</h3>
          <div style={{ height: 400 }}>
            <ResponsiveContainer>
              <BarChart data={consumoPorCategoria}>
                <CartesianGrid stroke="#2a2a3e" strokeDasharray="3 3" />
                <XAxis dataKey="name" stroke="#94a3b8" fontSize={12} />
                <YAxis stroke="#94a3b8" fontSize={12} />
                <Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} />
                <Legend />
                <Bar dataKey="stock" name="Stock Actual" fill="#22c55e" radius={[8,8,0,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        
        {/* Gráfico de evolución 30 días */}
        <div className="mb-8">
          <h3 className="text-base font-semibold text-slate-300 mb-4 flex items-center gap-2">📈 Evolución últimos 30 días</h3>
          <div style={{ height: 400 }}>
            <ResponsiveContainer>
              <AreaChart data={ultimos30Dias}>
                <CartesianGrid stroke="#2a2a3e" strokeDasharray="3 3" />
                <XAxis dataKey="fecha" stroke="#94a3b8" fontSize={10} interval={5} />
                <YAxis stroke="#94a3b8" fontSize={12} />
                <Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} />
                <Legend />
                <Area type="monotone" dataKey="entradas" name="Entradas" stackId="1" stroke="#22c55e" fill="#22c55e" fillOpacity={0.3} />
                <Area type="monotone" dataKey="salidas" name="Salidas" stackId="1" stroke="#ef4444" fill="#ef4444" fillOpacity={0.3} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
        
        {/* Movimientos últimos 7 días */}
        <div className="mb-8">
          <h3 className="text-base font-semibold text-slate-300 mb-4 flex items-center gap-2">🔄 Movimientos últimos 7 días</h3>
          <div style={{ height: 400 }}>
            <ResponsiveContainer>
              <LineChart data={ultimos7Dias}>
                <CartesianGrid stroke="#2a2a3e" strokeDasharray="3 3" />
                <XAxis dataKey="fecha" stroke="#94a3b8" fontSize={12} />
                <YAxis stroke="#94a3b8" fontSize={12} />
                <Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} />
                <Legend />
                <Line type="monotone" dataKey="entradas" name="Entradas" stroke="#22c55e" strokeWidth={3} dot={{ r: 6, fill: "#22c55e" }} />
                <Line type="monotone" dataKey="salidas" name="Salidas" stroke="#ef4444" strokeWidth={3} dot={{ r: 6, fill: "#ef4444" }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
        
        {/* Top 10 productos más vendidos */}
        <div className="mb-8">
          <h3 className="text-base font-semibold text-slate-300 mb-4 flex items-center gap-2">🏆 Top 10 productos más vendidos</h3>
          <div style={{ height: 500 }}>
            <ResponsiveContainer>
              <BarChart data={topProductos} layout="vertical">
                <CartesianGrid stroke="#2a2a3e" strokeDasharray="3 3" />
                <XAxis type="number" stroke="#94a3b8" fontSize={12} />
                <YAxis type="category" dataKey="nombre" stroke="#94a3b8" fontSize={10} width={140} />
                <Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} />
                <Legend />
                <Bar dataKey="salidas" name="Unidades Vendidas" fill="#f97316" radius={[0,8,8,0]} />
                <Bar dataKey="stock" name="Stock Actual" fill="#22c55e" radius={[0,8,8,0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        
        {/* Distribución por categoría */}
        <div className="mb-6">
          <h3 className="text-base font-semibold text-slate-300 mb-4 flex items-center gap-2">🥧 Distribución de Stock por Categoría</h3>
          <div style={{ height: 450 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie 
                  data={consumoPorCategoria} 
                  cx="50%" 
                  cy="50%" 
                  labelLine={true} 
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`} 
                  outerRadius={160} 
                  fill="#8884d8" 
                  dataKey="stock"
                >
                  {consumoPorCategoria.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
        
        <div className="flex justify-end pt-4">
          <button onClick={onClose} className="px-6 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium transition-all">Cerrar</button>
        </div>
      </div>
    </div>
  );
}

function ModalReportes({ productos, movimientos, onClose }: { productos: Producto[]; movimientos: Movimiento[]; onClose: () => void }) {
  const [fechaInicio, setFechaInicio] = useState(new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]);
  const [fechaFin, setFechaFin] = useState(new Date().toISOString().split('T')[0]);
  const [tipoReporte, setTipoReporte] = useState<TipoReporte>("todos");
  const [productoFiltro, setProductoFiltro] = useState<string>("");
  const [exportando, setExportando] = useState(false);
  const [formatoExportacion, setFormatoExportacion] = useState<"excel" | "pdf">("excel");

  const movimientosFiltrados = useMemo(() => {
    return movimientos.filter(m => {
      const fechaMov = m.created_at?.split('T')[0];
      const cumpleFecha = fechaMov >= fechaInicio && fechaMov <= fechaFin;
      const cumpleTipo = tipoReporte === "todos" || m.tipo === tipoReporte;
      const cumpleProducto = !productoFiltro || m.producto_id === parseInt(productoFiltro);
      return cumpleFecha && cumpleTipo && cumpleProducto;
    });
  }, [movimientos, fechaInicio, fechaFin, tipoReporte, productoFiltro]);

  const exportarExcel = async () => {
    setExportando(true);
    try {
      const wb = new ExcelJS.Workbook();
      const sh = wb.addWorksheet("Reporte_Movimientos");
      
      // Estilos
      const headerStyle = {
        font: { bold: true, size: 12, color: { argb: "FFFFFFFF" } },
        fill: { type: "pattern" as const, pattern: "solid", fgColor: { argb: "FF1f2937" } }
      };
      
      const headerRow = sh.addRow(["Fecha", "Producto", "Categoría", "Tipo", "Cantidad", "Motivo", "Notas", "Usuario"]);
      headerRow.eachCell(cell => {
        cell.font = headerStyle.font;
        cell.fill = headerStyle.fill;
      });
      
      movimientosFiltrados.forEach(m => {
        const prod = productos.find(p => p.id === m.producto_id);
        sh.addRow([
          formatFecha(m.created_at), 
          prod?.nombre || "?", 
          prod?.categoria || "?", 
          m.tipo === "entrada" ? "Ingreso" : "Salida", 
          m.cantidad, 
          m.motivo || "-", 
          m.notas || "-",
          (m as any).usuario || "-"
        ]);
      });
      
      sh.columns = [{ width: 20 }, { width: 35 }, { width: 18 }, { width: 12 }, { width: 12 }, { width: 25 }, { width: 35 }, { width: 15 }];
      
      // Resumen
      sh.addRow([]);
      const resumenRow = sh.addRow(["RESUMEN DEL PERÍODO"]);
      resumenRow.font = { bold: true, size: 14 };
      sh.addRow([`Total movimientos: ${movimientosFiltrados.length}`]);
      sh.addRow([`Total entradas: ${movimientosFiltrados.filter(m => m.tipo === "entrada").reduce((s, m) => s + m.cantidad, 0)} unidades`]);
      sh.addRow([`Total salidas: ${movimientosFiltrados.filter(m => m.tipo === "salida").reduce((s, m) => s + m.cantidad, 0)} unidades`]);
      sh.addRow([`Período: ${fechaInicio} al ${fechaFin}`]);
      sh.addRow([`Generado: ${new Date().toLocaleString("es-CO")}`]);
      
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `reporte_inventario_${fechaInicio}_a_${fechaFin}.xlsx`;
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Error al exportar:", error);
      alert("Error al generar el reporte");
    } finally {
      setExportando(false);
    }
  };

  const exportarPDF = async () => {
    setExportando(true);
    try {
      const doc = new jsPDF();
      
      // Título
      doc.setFontSize(18);
      doc.setTextColor(251, 191, 36);
      doc.text("Reporte de Movimientos de Inventario", 14, 20);
      
      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.text(`Período: ${fechaInicio} al ${fechaFin}`, 14, 30);
      doc.text(`Generado: ${new Date().toLocaleString("es-CO")}`, 14, 36);
      
      // Resumen
      doc.setFontSize(12);
      doc.setTextColor(0);
      doc.text("Resumen:", 14, 48);
      doc.setFontSize(10);
      doc.text(`Total movimientos: ${movimientosFiltrados.length}`, 20, 56);
      doc.text(`Total entradas: ${movimientosFiltrados.filter(m => m.tipo === "entrada").reduce((s, m) => s + m.cantidad, 0)} unidades`, 20, 62);
      doc.text(`Total salidas: ${movimientosFiltrados.filter(m => m.tipo === "salida").reduce((s, m) => s + m.cantidad, 0)} unidades`, 20, 68);
      
      // Tabla de movimientos
      const tableData = movimientosFiltrados.slice(0, 100).map(m => {
        const prod = productos.find(p => p.id === m.producto_id);
        return [
          formatFecha(m.created_at),
          prod?.nombre || "?",
          m.tipo === "entrada" ? "Ingreso" : "Salida",
          m.cantidad.toString(),
          m.motivo || "-"
        ];
      });
      
      autoTable(doc, {
        startY: 80,
        head: [["Fecha", "Producto", "Tipo", "Cantidad", "Motivo"]],
        body: tableData,
        theme: "striped",
        headStyles: { fillColor: [251, 191, 36], textColor: [0, 0, 0], fontStyle: "bold" },
        styles: { fontSize: 8, cellPadding: 2 },
        margin: { left: 14, right: 14 }
      });
      
      if (movimientosFiltrados.length > 100) {
        doc.text(`* Mostrando los primeros 100 de ${movimientosFiltrados.length} movimientos`, 14, (doc as any).lastAutoTable.finalY + 10);
      }
      
      doc.save(`reporte_inventario_${fechaInicio}_a_${fechaFin}.pdf`);
    } catch (error) {
      console.error("Error al exportar PDF:", error);
      alert("Error al generar el PDF");
    } finally {
      setExportando(false);
    }
  };

  const handleExportar = () => {
    if (formatoExportacion === "excel") {
      exportarExcel();
    } else {
      exportarPDF();
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6 animate-in fade-in duration-200" onClick={onClose}>
      <div className="relative w-full max-w-4xl max-h-[90vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl animate-in zoom-in-95 duration-300" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-6 sticky top-0 bg-[#0f1117] z-10 pb-4 border-b border-[#2a2a3e]">
          <h2 className="text-2xl font-bold bg-gradient-to-r from-amber-400 to-amber-600 bg-clip-text text-transparent">📋 Generar Reporte Avanzado</h2>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>
        
        <div className="grid grid-cols-3 gap-5 mb-6">
          <div>
            <label className="block text-sm font-medium text-slate-400 mb-2">Fecha Inicio</label>
            <input type="date" value={fechaInicio} onChange={(e) => setFechaInicio(e.target.value)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20" />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-400 mb-2">Fecha Fin</label>
            <input type="date" value={fechaFin} onChange={(e) => setFechaFin(e.target.value)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20" />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-400 mb-2">Tipo de Movimiento</label>
            <select value={tipoReporte} onChange={(e) => setTipoReporte(e.target.value as TipoReporte)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500">
              <option value="todos">Todos</option>
              <option value="entradas">Solo Entradas</option>
              <option value="salidas">Solo Salidas</option>
            </select>
          </div>
        </div>
        
        <div className="mb-6">
          <label className="block text-sm font-medium text-slate-400 mb-2">Filtrar por Producto (opcional)</label>
          <select value={productoFiltro} onChange={(e) => setProductoFiltro(e.target.value)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500">
            <option value="">Todos los productos</option>
            {productos.map(p => <option key={p.id} value={p.id}>{p.nombre}</option>)}
          </select>
        </div>
        
        <div className="mb-6">
          <label className="block text-sm font-medium text-slate-400 mb-2">Formato de Exportación</label>
          <div className="flex gap-3">
            <button 
              onClick={() => setFormatoExportacion("excel")}
              className={`flex-1 py-3 rounded-xl text-sm font-semibold transition-all flex items-center justify-center gap-2 ${formatoExportacion === "excel" ? 'bg-gradient-to-r from-green-600 to-green-500 text-white shadow-md' : 'bg-[#0a0d16] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}
            >
              📊 Excel
            </button>
            <button 
              onClick={() => setFormatoExportacion("pdf")}
              className={`flex-1 py-3 rounded-xl text-sm font-semibold transition-all flex items-center justify-center gap-2 ${formatoExportacion === "pdf" ? 'bg-gradient-to-r from-red-600 to-red-500 text-white shadow-md' : 'bg-[#0a0d16] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}
            >
              📄 PDF
            </button>
          </div>
        </div>
        
        <div className="mb-6 p-4 rounded-xl bg-gradient-to-r from-amber-500/10 to-orange-500/10 border border-amber-500/30">
          <p className="text-sm text-amber-300">
            📊 Resumen del reporte: {movimientosFiltrados.length} movimientos encontrados | 
            ↑ Entradas: {movimientosFiltrados.filter(m => m.tipo === "entrada").reduce((s, m) => s + m.cantidad, 0)} unidades | 
            ↓ Salidas: {movimientosFiltrados.filter(m => m.tipo === "salida").reduce((s, m) => s + m.cantidad, 0)} unidades
          </p>
        </div>
        
        <div className="flex gap-4 justify-end pt-4">
          <button onClick={onClose} className="px-6 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium transition-all">Cancelar</button>
          <button onClick={handleExportar} disabled={exportando} className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white font-medium transition-all flex items-center gap-2 shadow-lg">
            {exportando ? (
              <>
                <div className="animate-spin rounded-full h-4 w-4 border-2 border-white border-t-transparent"></div>
                Exportando...
              </>
            ) : (
              <>📥 Exportar a {formatoExportacion.toUpperCase()}</>
            )}
          </button>
        </div>
      </div>
    </div>
  );
}

// ==================== MODAL ADMIN PRODUCTOS ====================

type ModalAdminProductosProps = {
  onClose: () => void;
  productos: Producto[];
  onProductoCreado: () => void;
  mostrarNotificacion: (tipo: Notificacion['tipo'], mensaje: string) => void;
};

function ModalAdminProductos({ onClose, productos, onProductoCreado, mostrarNotificacion }: ModalAdminProductosProps) {
  const [modo, setModo] = useState<"agregar" | "eliminar" | "editar">("agregar");
  const [nombre, setNombre] = useState("");
  const [categoria, setCategoria] = useState<Categoria>(Categoria.Cervezas);
  const [stock, setStock] = useState("");
  const [stockMinimo, setStockMinimo] = useState("");
  const [cantidadBase, setCantidadBase] = useState("");
  const [unidadesPorCaja, setUnidadesPorCaja] = useState(""); // NUEVO
  const [ubicacion, setUbicacion] = useState("");
  const [cargando, setCargando] = useState(false);
  const [productoAEliminar, setProductoAEliminar] = useState<number | null>(null);
  const [productoAEditar, setProductoAEditar] = useState<Producto | null>(null);
  const [errorDetalle, setErrorDetalle] = useState<string>("");

  const handleAgregar = async () => {
    setErrorDetalle("");
    
    if (!nombre.trim()) {
      setErrorDetalle("❌ El nombre del producto es obligatorio");
      return;
    }
    
    const stockValid = validateCantidad(stock);
    if (!stockValid.valida) {
      setErrorDetalle(`❌ ${stockValid.error}`);
      return;
    }
    
    const stockMinimoValid = validateCantidad(stockMinimo);
    if (!stockMinimoValid.valida) {
      setErrorDetalle(`❌ Stock mínimo: ${stockMinimoValid.error}`);
      return;
    }
    
    const cantidadBaseValid = validateCantidad(cantidadBase);
    if (!cantidadBaseValid.valida) {
      setErrorDetalle(`❌ Cantidad base: ${cantidadBaseValid.error}`);
      return;
    }
    
    const unidadesPorCajaValid = validateCantidad(unidadesPorCaja);
    if (!unidadesPorCajaValid.valida) {
      setErrorDetalle(`❌ Unidades por caja: ${unidadesPorCajaValid.error}`);
      return;
    }

    setCargando(true);
    
    try {
      const nuevoProducto = {
        nombre: nombre.trim(),
        categoria: categoria,
        stock: parseInt(stock),
        stock_minimo: parseInt(stockMinimo),
        cantidad_base: parseInt(cantidadBase),
        unidades_por_caja: parseInt(unidadesPorCaja) || 1,
        ubicacion: ubicacion || null,
      };
      
      const { error } = await supabase
        .from("productos")
        .insert(nuevoProducto);
      
      if (error) {
        console.error("Error de Supabase:", error);
        
        if (error.code === "23505") {
          setErrorDetalle("❌ Ya existe un producto con este nombre");
        } else {
          setErrorDetalle(`❌ Error: ${error.message}`);
        }
      } else {
        mostrarNotificacion('success', `✅ Producto "${nombre}" agregado exitosamente`);
        onProductoCreado();
        setNombre("");
        setCategoria(Categoria.Cervezas);
        setStock("");
        setStockMinimo("");
        setCantidadBase("");
        setUnidadesPorCaja("");
        setUbicacion("");
        setErrorDetalle("");
      }
    } catch (err: any) {
      console.error("Error inesperado:", err);
      setErrorDetalle(`❌ Error inesperado: ${err.message || "Desconocido"}`);
    } finally {
      setCargando(false);
    }
  };

  const handleEditar = async () => {
    if (!productoAEditar) return;
    
    setCargando(true);
    setErrorDetalle("");
    
    try {
      const updates: Partial<Producto> = {};
      if (nombre && nombre !== productoAEditar.nombre) updates.nombre = nombre;
      if (categoria !== productoAEditar.categoria) updates.categoria = categoria;
      if (stock) updates.stock = parseInt(stock);
      if (stockMinimo) updates.stock_minimo = parseInt(stockMinimo);
      if (cantidadBase) updates.cantidad_base = parseInt(cantidadBase);
      if (unidadesPorCaja) updates.unidades_por_caja = parseInt(unidadesPorCaja);
      if (ubicacion !== productoAEditar.ubicacion) updates.ubicacion = ubicacion || null;
      
      if (Object.keys(updates).length === 0) {
        setErrorDetalle("❌ No se realizaron cambios");
        setCargando(false);
        return;
      }
      
      const { error } = await supabase
        .from("productos")
        .update(updates)
        .eq("id", productoAEditar.id);
      
      if (error) {
        setErrorDetalle(`❌ Error al editar: ${error.message}`);
      } else {
        mostrarNotificacion('success', `✅ Producto "${productoAEditar.nombre}" actualizado correctamente`);
        onProductoCreado();
        setModo("agregar");
        setProductoAEditar(null);
        limpiarFormulario();
      }
    } catch (err: any) {
      setErrorDetalle(`❌ Error inesperado: ${err.message}`);
    } finally {
      setCargando(false);
    }
  };

  const handleEliminar = async () => {
    if (!productoAEliminar) return;

    const producto = productos.find(p => p.id === productoAEliminar);
    if (!confirm(`⚠️ ¿Estás seguro de eliminar "${producto?.nombre}"?\n\nEsta acción también eliminará todos sus movimientos asociados.`)) return;

    setCargando(true);
    setErrorDetalle("");
    
    try {
      const { error: errorMovimientos } = await supabase
        .from("movimientos")
        .delete()
        .eq("producto_id", productoAEliminar);
      
      if (errorMovimientos) {
        console.error("Error al eliminar movimientos:", errorMovimientos);
      }
      
      const { error: errorProducto } = await supabase
        .from("productos")
        .delete()
        .eq("id", productoAEliminar);
      
      if (errorProducto) {
        console.error("Error al eliminar producto:", errorProducto);
        setErrorDetalle(`❌ Error al eliminar: ${errorProducto.message}`);
      } else {
        mostrarNotificacion('success', `🗑️ Producto "${producto?.nombre}" eliminado correctamente`);
        onProductoCreado();
        setProductoAEliminar(null);
      }
    } catch (err: any) {
      console.error("Error inesperado:", err);
      setErrorDetalle(`❌ Error inesperado: ${err.message || "Desconocido"}`);
    } finally {
      setCargando(false);
    }
  };

  const limpiarFormulario = () => {
    setNombre("");
    setCategoria(Categoria.Cervezas);
    setStock("");
    setStockMinimo("");
    setCantidadBase("");
    setUnidadesPorCaja("");
    setUbicacion("");
    setErrorDetalle("");
  };

  const seleccionarProductoEditar = (producto: Producto) => {
    setProductoAEditar(producto);
    setNombre(producto.nombre);
    setCategoria(producto.categoria);
    setStock(producto.stock.toString());
    setStockMinimo(producto.stock_minimo.toString());
    setCantidadBase(producto.cantidad_base.toString());
    setUnidadesPorCaja(producto.unidades_por_caja?.toString() || "1");
    setUbicacion(producto.ubicacion || "");
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6 animate-in fade-in duration-200" onClick={onClose}>
      <div className="relative w-full max-w-2xl max-h-[90vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] shadow-2xl animate-in zoom-in-95 duration-300" onClick={(e) => e.stopPropagation()}>
        <div className="sticky top-0 bg-[#0f1117] border-b border-[#2a2a3e] p-6 flex justify-between items-center z-10">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center text-2xl bg-amber-500/20">🛠️</div>
            <div>
              <h2 className="text-xl font-bold text-white">Administrar Productos</h2>
              <p className="text-sm text-slate-500">Agrega, edita o elimina productos del inventario</p>
            </div>
          </div>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>

        {errorDetalle && (
          <div className="mx-6 mt-4 p-3 rounded-xl bg-red-500/10 border border-red-500/30 animate-shake">
            <p className="text-sm text-red-400">{errorDetalle}</p>
          </div>
        )}

        <div className="flex mx-6 mt-6 bg-[#0a0d16] rounded-xl p-1 gap-1">
          <button 
            onClick={() => { setModo("agregar"); limpiarFormulario(); setErrorDetalle(""); }} 
            className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${modo === "agregar" ? 'bg-gradient-to-r from-green-600 to-green-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}
          >
            ➕ Agregar
          </button>
          <button 
            onClick={() => { setModo("editar"); limpiarFormulario(); setErrorDetalle(""); }} 
            className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${modo === "editar" ? 'bg-gradient-to-r from-blue-600 to-blue-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}
          >
            ✏️ Editar
          </button>
          <button 
            onClick={() => { setModo("eliminar"); setErrorDetalle(""); }} 
            className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${modo === "eliminar" ? 'bg-gradient-to-r from-red-600 to-red-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}
          >
            🗑️ Eliminar
          </button>
        </div>

        <div className="p-6">
          {modo === "agregar" ? (
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-400 mb-2">Nombre del Producto *</label>
                <input 
                  type="text" 
                  value={nombre} 
                  onChange={(e) => setNombre(e.target.value)} 
                  placeholder="Ej: Corona Extra, Coca-Cola, Agua Mineral" 
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20" 
                  autoFocus
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-slate-400 mb-2">Categoría *</label>
                <select 
                  value={categoria} 
                  onChange={(e) => setCategoria(e.target.value as Categoria)} 
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"
                >
                  {CATEGORIAS.map(cat => (
                    <option key={cat} value={cat}>
                      {CATEGORIA_EMOJI[cat]} {cat}
                    </option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-slate-400 mb-2">Stock Actual *</label>
                  <input 
                    type="number" 
                    value={stock} 
                    onChange={(e) => setStock(e.target.value)} 
                    placeholder="0" 
                    className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" 
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-400 mb-2">Stock Mínimo *</label>
                  <input 
                    type="number" 
                    value={stockMinimo} 
                    onChange={(e) => setStockMinimo(e.target.value)} 
                    placeholder="Ej: 10" 
                    className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" 
                  />
                </div>
              </div>

              <div>
                <label className="block text-sm font-medium text-slate-400 mb-2">
                  Cantidad Base (Objetivo) *
                </label>
                <input 
                  type="number" 
                  value={cantidadBase} 
                  onChange={(e) => setCantidadBase(e.target.value)} 
                  placeholder="Cantidad que debe haber en el bar" 
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" 
                />
                <p className="text-xs text-slate-500 mt-1">Ej: Si siempre necesitas 100 cervezas, pon 100</p>
              </div>

              {/* NUEVO CAMPO: Unidades por Caja */}
              <div>
                <label className="block text-sm font-medium text-slate-400 mb-2">
                  📦 Unidades por Caja / Paca *
                </label>
                <input 
                  type="number" 
                  value={unidadesPorCaja} 
                  onChange={(e) => setUnidadesPorCaja(e.target.value)} 
                  placeholder="Ej: 12, 6, 24" 
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" 
                />
                <p className="text-xs text-slate-500 mt-1">
                  ¿Cuántas unidades vienen en una caja/paca? Ej: Coca-Cola = 12, Agua = 24, Cerveza = 6, Soda = 6
                </p>
              </div>

              <div>
                <label className="block text-sm font-medium text-slate-400 mb-2">Ubicación</label>
                <input 
                  type="text" 
                  value={ubicacion} 
                  onChange={(e) => setUbicacion(e.target.value)} 
                  placeholder="Ej: Estante A1" 
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" 
                />
              </div>

              <div className="pt-4">
                <button 
                  onClick={handleAgregar} 
                  disabled={cargando || !nombre || !stock || !stockMinimo || !cantidadBase || !unidadesPorCaja} 
                  className="w-full py-3 rounded-xl bg-gradient-to-r from-green-600 to-green-500 hover:from-green-500 hover:to-green-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2 shadow-lg"
                >
                  {cargando ? (
                    <>
                      <div className="animate-spin rounded-full h-5 w-5 border-2 border-white border-t-transparent"></div>
                      AGREGANDO...
                    </>
                  ) : "➕ AGREGAR PRODUCTO"}
                </button>
              </div>
            </div>
          ) : modo === "editar" ? (
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-400 mb-2">Seleccionar producto a editar</label>
                <select 
                  onChange={(e) => {
                    const prod = productos.find(p => p.id === parseInt(e.target.value));
                    if (prod) seleccionarProductoEditar(prod);
                  }}
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"
                >
                  <option value="">-- Seleccionar producto --</option>
                  {productos.map(p => (
                    <option key={p.id} value={p.id}>
                      {p.nombre} - Stock: {p.stock} und
                    </option>
                  ))}
                </select>
              </div>

              {productoAEditar && (
                <>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Nombre</label>
                    <input 
                      type="text" 
                      value={nombre} 
                      onChange={(e) => setNombre(e.target.value)} 
                      className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm" 
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Categoría</label>
                    <select 
                      value={categoria} 
                      onChange={(e) => setCategoria(e.target.value as Categoria)} 
                      className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm"
                    >
                      {CATEGORIAS.map(cat => (
                        <option key={cat} value={cat}>{cat}</option>
                      ))}
                    </select>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-sm font-medium text-slate-400 mb-2">Stock</label>
                      <input 
                        type="number" 
                        value={stock} 
                        onChange={(e) => setStock(e.target.value)} 
                        className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm" 
                      />
                    </div>
                    <div>
                      <label className="block text-sm font-medium text-slate-400 mb-2">Stock Mínimo</label>
                      <input 
                        type="number" 
                        value={stockMinimo} 
                        onChange={(e) => setStockMinimo(e.target.value)} 
                        className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm" 
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Cantidad Base</label>
                    <input 
                      type="number" 
                      value={cantidadBase} 
                      onChange={(e) => setCantidadBase(e.target.value)} 
                      className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm" 
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">📦 Unidades por Caja</label>
                    <input 
                      type="number" 
                      value={unidadesPorCaja} 
                      onChange={(e) => setUnidadesPorCaja(e.target.value)} 
                      className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm" 
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Ubicación</label>
                    <input 
                      type="text" 
                      value={ubicacion} 
                      onChange={(e) => setUbicacion(e.target.value)} 
                      className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm" 
                    />
                  </div>
                  <button 
                    onClick={handleEditar} 
                    disabled={cargando} 
                    className="w-full py-3 rounded-xl bg-gradient-to-r from-blue-600 to-blue-500 hover:from-blue-500 hover:to-blue-600 text-white font-semibold transition-all flex items-center justify-center gap-2"
                  >
                    {cargando ? "GUARDANDO..." : "💾 GUARDAR CAMBIOS"}
                  </button>
                </>
              )}
            </div>
          ) : (
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-slate-400 mb-2">Seleccionar producto a eliminar</label>
                <select 
                  value={productoAEliminar || ""} 
                  onChange={(e) => setProductoAEliminar(parseInt(e.target.value))} 
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"
                >
                  <option value="">-- Seleccionar producto --</option>
                  {productos.map(p => (
                    <option key={p.id} value={p.id}>
                      {p.nombre} - Stock: {p.stock} und - Base: {p.cantidad_base} und - Caja: {p.unidades_por_caja || 1} unid
                    </option>
                  ))}
                </select>
              </div>

              {productoAEliminar && (
                <div className="p-4 rounded-xl bg-red-500/10 border border-red-500/30">
                  <p className="text-sm text-red-300 mb-2">⚠️ Advertencia:</p>
                  <p className="text-xs text-slate-400">
                    Se eliminará el producto y <strong className="text-red-400">todos sus movimientos asociados</strong>. 
                    Esta acción no se puede deshacer.
                  </p>
                </div>
              )}

              <button 
                onClick={handleEliminar} 
                disabled={cargando || !productoAEliminar} 
                className="w-full py-3 rounded-xl bg-gradient-to-r from-red-600 to-red-500 hover:from-red-500 hover:to-red-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {cargando ? (
                  <>
                    <div className="animate-spin rounded-full h-5 w-5 border-2 border-white border-t-transparent"></div>
                    ELIMINANDO...
                  </>
                ) : "🗑️ ELIMINAR PRODUCTO"}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ==================== MODAL EDITAR CANTIDAD BASE (MEJORADO) ====================

function ModalEditarCantidadBase({ producto, onClose, onActualizado, mostrarNotificacion }: { producto: Producto; onClose: () => void; onActualizado: () => void; mostrarNotificacion: (tipo: Notificacion['tipo'], mensaje: string) => void }) {
  const [cantidadBase, setCantidadBase] = useState(producto.cantidad_base?.toString() || "");
  const [unidadesPorCaja, setUnidadesPorCaja] = useState(producto.unidades_por_caja?.toString() || "1");
  const [cargando, setCargando] = useState(false);

  const handleGuardar = async () => {
    const nuevaCantidad = parseInt(cantidadBase);
    const nuevasUnidadesPorCaja = parseInt(unidadesPorCaja);
    
    if (isNaN(nuevaCantidad) || nuevaCantidad < 0) {
      mostrarNotificacion('error', "Ingrese una cantidad base válida");
      return;
    }
    
    if (isNaN(nuevasUnidadesPorCaja) || nuevasUnidadesPorCaja < 1) {
      mostrarNotificacion('error', "Ingrese unidades por caja válidas (mínimo 1)");
      return;
    }

    setCargando(true);
    const { error } = await supabase
      .from("productos")
      .update({ 
        cantidad_base: nuevaCantidad,
        unidades_por_caja: nuevasUnidadesPorCaja 
      })
      .eq("id", producto.id);

    if (error) {
      console.error("Error:", error);
      mostrarNotificacion('error', "Error al actualizar");
    } else {
      mostrarNotificacion('success', `✅ "${producto.nombre}" actualizado: Base ${nuevaCantidad}, Caja ${nuevasUnidadesPorCaja} unid`);
      onActualizado();
      onClose();
    }
    setCargando(false);
  };

  const pedidoUnidades = calcularPedidoSugerido(producto.stock, parseInt(cantidadBase) || 0);
  const { cajas, resto } = calcularCajasNecesarias(pedidoUnidades, parseInt(unidadesPorCaja) || 1);
  const textoPedido = cajas > 0 && resto > 0 ? `${cajas} caja${cajas !== 1 ? 's' : ''} + ${resto} und` 
    : cajas > 0 ? `${cajas} caja${cajas !== 1 ? 's' : ''}` 
    : `${resto} unidades`;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6 animate-in fade-in duration-200" onClick={onClose}>
      <div className="relative w-full max-w-md rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl animate-in zoom-in-95 duration-300" onClick={(e) => e.stopPropagation()}>
        <h2 className="text-xl font-bold text-white mb-4">Editar Producto</h2>
        <p className="text-slate-400 mb-4">Producto: <span className="text-white font-semibold">{producto.nombre}</span></p>
        
        <div className="mb-4">
          <label className="block text-sm font-medium text-slate-400 mb-2">Cantidad Base (Objetivo)</label>
          <input 
            type="number" 
            value={cantidadBase} 
            onChange={(e) => setCantidadBase(e.target.value)} 
            className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-lg text-center focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20"
            autoFocus
          />
          <p className="text-xs text-slate-500 mt-2">Cantidad que siempre debe haber en el bar</p>
        </div>

        <div className="mb-4">
          <label className="block text-sm font-medium text-slate-400 mb-2">📦 Unidades por Caja</label>
          <input 
            type="number" 
            value={unidadesPorCaja} 
            onChange={(e) => setUnidadesPorCaja(e.target.value)} 
            className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-lg text-center focus:outline-none focus:border-amber-500"
          />
          <p className="text-xs text-slate-500 mt-2">¿Cuántas unidades vienen en una caja?</p>
        </div>

        <div className="bg-[#0a0d16] rounded-xl p-3 mb-4">
          <p className="text-xs text-slate-400">Stock actual: <span className="text-white font-bold">{producto.stock}</span> unidades</p>
          <p className="text-xs text-orange-400 mt-1">Pedido sugerido: {textoPedido} ({pedidoUnidades} und)</p>
        </div>

        <div className="flex gap-3">
          <button onClick={onClose} className="flex-1 py-3 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium transition-all">Cancelar</button>
          <button onClick={handleGuardar} disabled={cargando} className="flex-1 py-3 rounded-xl bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white font-semibold transition-all shadow-lg">
            {cargando ? "Guardando..." : "Guardar"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ==================== MODAL RESTAURAR BACKUP ====================

function ModalRestaurarBackup({ onClose, onRestaurar, mostrarNotificacion }: { onClose: () => void; onRestaurar: (productos: Producto[], movimientos: Movimiento[]) => void; mostrarNotificacion: (tipo: Notificacion['tipo'], mensaje: string) => void }) {
  const [backups, setBackups] = useState<Array<{ fecha: string; productos: Producto[]; movimientos: Movimiento[] }>>([]);

  useEffect(() => {
    const cargarBackups = () => {
      const backupsGuardados: Array<{ fecha: string; productos: Producto[]; movimientos: Movimiento[] }> = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key?.startsWith('backup_')) {
          try {
            const data = JSON.parse(localStorage.getItem(key) || '');
            if (data.productos && data.movimientos) {
              backupsGuardados.push({
                fecha: data.fecha || key.replace('backup_', ''),
                productos: data.productos,
                movimientos: data.movimientos
              });
            }
          } catch (e) {}
        }
      }
      backupsGuardados.sort((a, b) => new Date(b.fecha).getTime() - new Date(a.fecha).getTime());
      setBackups(backupsGuardados);
    };
    cargarBackups();
  }, []);

  const handleRestaurar = (backup: { productos: Producto[]; movimientos: Movimiento[] }) => {
    if (confirm("⚠️ Esta acción sobrescribirá los datos actuales. ¿Continuar?")) {
      onRestaurar(backup.productos, backup.movimientos);
      mostrarNotificacion('success', "Backup restaurado correctamente");
      onClose();
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6 animate-in fade-in duration-200" onClick={onClose}>
      <div className="relative w-full max-w-3xl max-h-[80vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl animate-in zoom-in-95 duration-300" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-6">
          <h2 className="text-xl font-bold text-white">💾 Restaurar Backup</h2>
          <button onClick={onClose} className="w-8 h-8 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-400">✕</button>
        </div>

        {backups.length === 0 ? (
          <div className="text-center py-12 text-slate-500">
            No hay backups disponibles. Los backups se crean automáticamente cada 5 minutos.
          </div>
        ) : (
          <div className="space-y-3">
            {backups.map((backup, idx) => (
              <div key={idx} className="flex items-center justify-between p-4 rounded-xl bg-[#0a0d16] border border-[#2a2a3e]">
                <div>
                  <div className="text-white font-medium">{new Date(backup.fecha).toLocaleString()}</div>
                  <div className="text-xs text-slate-500">
                    {backup.productos.length} productos | {backup.movimientos.length} movimientos
                  </div>
                </div>
                <button 
                  onClick={() => handleRestaurar(backup)}
                  className="px-4 py-2 rounded-lg bg-amber-500/20 text-amber-400 text-sm hover:bg-amber-500/30 transition-all"
                >
                  Restaurar
                </button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ==================== VISTA DE INVENTARIO ====================

function VistaInventario({ productos, onEditarCantidadBase }: { productos: Producto[]; onEditarCantidadBase: (producto: Producto) => void }) {
  const [busqueda, setBusqueda] = useState("");
  const [categoriaFiltro, setCategoriaFiltro] = useState<Categoria | null>(null);
  const [paginaActual, setPaginaActual] = useState(1);
  const [ordenarPor, setOrdenarPor] = useState<"nombre" | "stock" | "categoria">("nombre");
  const [ordenDireccion, setOrdenDireccion] = useState<"asc" | "desc">("asc");
  const itemsPorPagina = 20;

  const productosOrdenados = useMemo(() => {
    let sorted = [...productos];
    switch (ordenarPor) {
      case "nombre":
        sorted.sort((a, b) => ordenDireccion === "asc" ? a.nombre.localeCompare(b.nombre) : b.nombre.localeCompare(a.nombre));
        break;
      case "stock":
        sorted.sort((a, b) => ordenDireccion === "asc" ? a.stock - b.stock : b.stock - a.stock);
        break;
      case "categoria":
        sorted.sort((a, b) => ordenDireccion === "asc" ? a.categoria.localeCompare(b.categoria) : b.categoria.localeCompare(a.categoria));
        break;
    }
    return sorted;
  }, [productos, ordenarPor, ordenDireccion]);

  const productosFiltrados = useMemo(() => {
    return productosOrdenados.filter(p => 
      p.nombre.toLowerCase().includes(busqueda.toLowerCase()) && 
      (!categoriaFiltro || p.categoria === categoriaFiltro)
    );
  }, [productosOrdenados, busqueda, categoriaFiltro]);

  const totalPaginas = Math.ceil(productosFiltrados.length / itemsPorPagina);
  const productosPaginados = productosFiltrados.slice(
    (paginaActual - 1) * itemsPorPagina,
    paginaActual * itemsPorPagina
  );

  useEffect(() => {
    setPaginaActual(1);
  }, [busqueda, categoriaFiltro, ordenarPor, ordenDireccion]);

  const totalPedidos = productos.reduce((sum, p) => sum + calcularPedidoSugerido(p.stock, p.cantidad_base), 0);

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold bg-gradient-to-r from-white to-slate-400 bg-clip-text text-transparent mb-2">📦 Inventario General</h2>
        <p className="text-slate-500">Listado completo de todos los productos con cantidades actuales y pedidos sugeridos</p>
      </div>

      <div className="flex gap-3 mb-6 flex-wrap items-center">
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">🔍</span>
          <input 
            type="text" 
            placeholder="Buscar producto..." 
            value={busqueda} 
            onChange={(e) => setBusqueda(e.target.value)} 
            className="pl-9 pr-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-72 focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20 transition-all" 
          />
        </div>
        
        <select 
          value={ordenarPor} 
          onChange={(e) => setOrdenarPor(e.target.value as any)}
          className="px-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm"
        >
          <option value="nombre">Ordenar por nombre</option>
          <option value="stock">Ordenar por stock</option>
          <option value="categoria">Ordenar por categoría</option>
        </select>
        
        <button 
          onClick={() => setOrdenDireccion(ordenDireccion === "asc" ? "desc" : "asc")}
          className="px-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm"
        >
          {ordenDireccion === "asc" ? "↑ Ascendente" : "↓ Descendente"}
        </button>
        
        <button onClick={() => setCategoriaFiltro(null)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all ${!categoriaFiltro ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>Todas</button>
        {CATEGORIAS.map(cat => (
          <button key={cat} onClick={() => setCategoriaFiltro(cat)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all flex items-center gap-2 ${categoriaFiltro === cat ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>
            <span>{CATEGORIA_EMOJI[cat]}</span>{cat}
          </button>
        ))}
      </div>

      <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
        <div className="px-5 py-4 border-b border-[#2a2a3e] bg-[#0a0d16] flex justify-between items-center flex-wrap gap-2">
          <span className="text-base font-semibold text-white">📋 Lista de Productos ({productosFiltrados.length} productos)</span>
          <div className="text-sm">
            <span className="text-slate-400">Total a pedir: </span>
            <span className="text-orange-400 font-bold text-lg">{totalPedidos}</span>
            <span className="text-slate-400"> unidades</span>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-[#0a0d16] border-b border-[#2a2a3e]">
              <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wider">
                <th className="p-4">Producto</th>
                <th className="p-4 text-center">Stock</th>
                <th className="p-4 text-center">Base</th>
                <th className="p-4 text-center">Unid./Caja</th>
                <th className="p-4 text-center">Pedido</th>
                <th className="p-4">Estado</th>
                <th className="p-4">Ubicación</th>
                <th className="p-4">Acciones</th>
              </tr>
            </thead>
            <tbody>
              {productosPaginados.map((p, i) => {
                const est = estadoInfo(p);
                const pedidoSugerido = calcularPedidoSugerido(p.stock, p.cantidad_base);
                const { cajas, resto } = calcularCajasNecesarias(pedidoSugerido, p.unidades_por_caja || 1);
                const textoPedido = cajas > 0 ? `${cajas} caja${cajas !== 1 ? 's' : ''}${resto > 0 ? ` + ${resto}` : ''}` : `${resto} und`;
                return (
                  <tr key={p.id} className={`border-b border-[#1a1a2a] hover:bg-[#1a1a2a] transition-all ${i % 2 === 0 ? 'bg-transparent' : 'bg-[#0a0d16]/50'}`}>
                    <td className="p-4">
                      <div>
                        <span className="text-white font-medium text-base">{p.nombre}</span>
                        <span className="text-xs ml-2 px-2 py-1 rounded-full" style={{ background: CATEGORIA_BG[p.categoria], color: CATEGORIA_COLOR[p.categoria], border: `1px solid ${CATEGORIA_BORDER[p.categoria]}` }}>
                          {CATEGORIA_EMOJI[p.categoria]} {p.categoria}
                        </span>
                      </div>
                    </td>
                    <td className="p-4 text-center"><span className="text-3xl font-bold" style={{ color: est.color }}>{p.stock}</span><span className="text-sm text-slate-500 ml-1">und</span></td>
                    <td className="p-4 text-center"><span className="text-xl font-semibold text-blue-400">{p.cantidad_base}</span><span className="text-sm text-slate-500 ml-1">und</span></td>
                    <td className="p-4 text-center"><span className="text-lg font-semibold text-purple-400">{p.unidades_por_caja || 1}</span><span className="text-sm text-slate-500 ml-1">und/caja</span></td>
                    <td className="p-4 text-center">
                      {pedidoSugerido > 0 ? (
                        <div className="flex flex-col items-center">
                          <span className="text-2xl font-bold text-orange-400">{textoPedido}</span>
                          <span className="text-xs text-orange-500">({pedidoSugerido} und)</span>
                        </div>
                      ) : (
                        <span className="text-sm text-green-500">✓ Suficiente</span>
                      )}
                    </td>
                    <td className="p-4"><span className="text-xs px-3 py-1.5 rounded-full font-semibold" style={{ background: est.bg, color: est.color, border: `1px solid ${est.color}40` }}><span className="inline-block w-2 h-2 rounded-full mr-1.5" style={{ background: est.dot }} />{est.label}</span></td>
                    <td className="p-4 text-slate-400 text-sm">{p.ubicacion || "—"}</td>
                    <td className="p-4">
                      <button 
                        onClick={() => onEditarCantidadBase(p)}
                        className="px-3 py-1.5 rounded-lg bg-amber-500/10 border border-amber-500/30 text-amber-400 text-xs font-medium hover:bg-amber-500/20 transition-all"
                      >
                        Editar
                      </button>
                    </td>
                  </tr>
                );
              })}
              {productosFiltrados.length === 0 && (
                <tr>
                  <td colSpan={8} className="text-center py-12 text-slate-500">No se encontraron productos</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        
        {totalPaginas > 1 && (
          <div className="flex justify-center gap-2 p-4 border-t border-[#2a2a3e]">
            <button 
              onClick={() => setPaginaActual(Math.max(1, paginaActual - 1))}
              disabled={paginaActual === 1}
              className="px-4 py-2 rounded-lg bg-slate-800 text-white text-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-700 transition-all"
            >
              Anterior
            </button>
            <span className="px-4 py-2 text-slate-400 text-sm">
              Página {paginaActual} de {totalPaginas}
            </span>
            <button 
              onClick={() => setPaginaActual(Math.min(totalPaginas, paginaActual + 1))}
              disabled={paginaActual === totalPaginas}
              className="px-4 py-2 rounded-lg bg-slate-800 text-white text-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-700 transition-all"
            >
              Siguiente
            </button>
          </div>
        )}
      </div>

      <div className="mt-6 grid grid-cols-5 gap-4">
        <div className="bg-gradient-to-br from-blue-500/10 to-blue-600/5 rounded-xl p-4 text-center border border-blue-500/30">
          <div className="text-2xl font-bold text-blue-400">{productos.length}</div>
          <div className="text-xs text-slate-500">Total Productos</div>
        </div>
        <div className="bg-gradient-to-br from-green-500/10 to-green-600/5 rounded-xl p-4 text-center border border-green-500/30">
          <div className="text-2xl font-bold text-green-400">{productos.reduce((sum, p) => sum + p.stock, 0)}</div>
          <div className="text-xs text-slate-500">Unidades Totales</div>
        </div>
        <div className="bg-gradient-to-br from-orange-500/10 to-orange-600/5 rounded-xl p-4 text-center border border-orange-500/30">
          <div className="text-2xl font-bold text-orange-400">{totalPedidos}</div>
          <div className="text-xs text-slate-500">Unidades a Pedir</div>
        </div>
        <div className="bg-gradient-to-br from-yellow-500/10 to-yellow-600/5 rounded-xl p-4 text-center border border-yellow-500/30">
          <div className="text-2xl font-bold text-yellow-400">{productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length}</div>
          <div className="text-xs text-slate-500">Stock Crítico</div>
        </div>
        <div className="bg-gradient-to-br from-red-500/10 to-red-600/5 rounded-xl p-4 text-center border border-red-500/30">
          <div className="text-2xl font-bold text-red-400">{productos.filter(p => p.stock === 0).length}</div>
          <div className="text-xs text-slate-500">Productos Agotados</div>
        </div>
      </div>
    </div>
  );
}

// ==================== VISTA DE MOVIMIENTOS ====================

function VistaMovimientos({ movimientos, productos }: { movimientos: Movimiento[]; productos: Producto[] }) {
  const [filtroTipo, setFiltroTipo] = useState<TipoReporte>("todos");
  const [filtroProducto, setFiltroProducto] = useState<string>("");
  const [fechaInicio, setFechaInicio] = useState("");
  const [fechaFin, setFechaFin] = useState("");
  const [paginaActual, setPaginaActual] = useState(1);
  const itemsPorPagina = 50;

  const movimientosFiltrados = useMemo(() => {
    return movimientos.filter(m => {
      const cumpleTipo = filtroTipo === "todos" || m.tipo === filtroTipo;
      const cumpleProducto = !filtroProducto || m.producto_id === parseInt(filtroProducto);
      const fechaMov = new Date(m.created_at).toISOString().split('T')[0];
      const cumpleFechaInicio = !fechaInicio || fechaMov >= fechaInicio;
      const cumpleFechaFin = !fechaFin || fechaMov <= fechaFin;
      return cumpleTipo && cumpleProducto && cumpleFechaInicio && cumpleFechaFin;
    });
  }, [movimientos, filtroTipo, filtroProducto, fechaInicio, fechaFin]);

  const totalPaginas = Math.ceil(movimientosFiltrados.length / itemsPorPagina);
  const movimientosPaginados = movimientosFiltrados.slice(
    (paginaActual - 1) * itemsPorPagina,
    paginaActual * itemsPorPagina
  );

  useEffect(() => {
    setPaginaActual(1);
  }, [filtroTipo, filtroProducto, fechaInicio, fechaFin]);

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold bg-gradient-to-r from-white to-slate-400 bg-clip-text text-transparent mb-2">🔄 Historial de Movimientos</h2>
        <p className="text-slate-500">Registro completo de todas las entradas y salidas de inventario</p>
      </div>

      <div className="grid grid-cols-4 gap-4 mb-6">
        <div>
          <label className="block text-xs font-medium text-slate-400 mb-1">Tipo de Movimiento</label>
          <select value={filtroTipo} onChange={(e) => setFiltroTipo(e.target.value as TipoReporte)} className="w-full px-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20">
            <option value="todos">Todos</option>
            <option value="entradas">↑ Entradas</option>
            <option value="salidas">↓ Salidas</option>
          </select>
        </div>
        <div>
          <label className="block text-xs font-medium text-slate-400 mb-1">Producto</label>
          <select value={filtroProducto} onChange={(e) => setFiltroProducto(e.target.value)} className="w-full px-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500">
            <option value="">Todos los productos</option>
            {productos.map(p => <option key={p.id} value={p.id}>{p.nombre}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-xs font-medium text-slate-400 mb-1">Fecha Desde</label>
          <input type="date" value={fechaInicio} onChange={(e) => setFechaInicio(e.target.value)} className="w-full px-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" />
        </div>
        <div>
          <label className="block text-xs font-medium text-slate-400 mb-1">Fecha Hasta</label>
          <input type="date" value={fechaFin} onChange={(e) => setFechaFin(e.target.value)} className="w-full px-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" />
        </div>
      </div>

      <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
        <div className="px-5 py-4 border-b border-[#2a2a3e] bg-[#0a0d16]">
          <span className="text-base font-semibold text-white">📋 Registro de Movimientos ({movimientosFiltrados.length} registros)</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-[#0a0d16] border-b border-[#2a2a3e]">
              <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wider">
                <th className="p-4">Fecha</th>
                <th className="p-4">Producto</th>
                <th className="p-4">Categoría</th>
                <th className="p-4">Tipo</th>
                <th className="p-4 text-center">Cantidad</th>
                <th className="p-4">Motivo</th>
                <th className="p-4">Notas</th>
               </tr>
            </thead>
            <tbody>
              {movimientosPaginados.map((m, i) => {
                const prod = productos.find(p => p.id === m.producto_id);
                const isEntrada = m.tipo === "entrada";
                return (
                  <tr key={m.id} className={`border-b border-[#1a1a2a] hover:bg-[#1a1a2a] transition-all ${i % 2 === 0 ? 'bg-transparent' : 'bg-[#0a0d16]/50'}`}>
                    <td className="p-4 text-slate-300 text-sm">{formatFecha(m.created_at)}</td>
                    <td className="p-4 text-white font-medium">{prod?.nombre || "?"}</td>
                    <td className="p-4"><span className="text-xs px-2 py-1 rounded-full" style={{ background: prod ? CATEGORIA_BG[prod.categoria] : '#1a1a2a', color: prod ? CATEGORIA_COLOR[prod.categoria] : '#94a3b8' }}>{prod ? CATEGORIA_EMOJI[prod.categoria] : '?'} {prod?.categoria || "?"}</span></td>
                    <td className="p-4"><span className={`text-xs px-3 py-1 rounded-full font-semibold ${isEntrada ? 'bg-green-500/20 text-green-400 border border-green-500/30' : 'bg-red-500/20 text-red-400 border border-red-500/30'}`}>{isEntrada ? "↑ ENTRADA" : "↓ SALIDA"}</span></td>
                    <td className="p-4 text-center"><span className={`text-xl font-bold ${isEntrada ? 'text-green-400' : 'text-red-400'}`}>{m.cantidad}</span></td>
                    <td className="p-4 text-slate-400 text-sm">{m.motivo || "—"}</td>
                    <td className="p-4 text-slate-500 text-sm">{m.notas || "—"}</td>
                  </tr>
                );
              })}
              {movimientosFiltrados.length === 0 && (
                <tr>
                  <td colSpan={7} className="text-center py-12 text-slate-500">No hay movimientos registrados</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        
        {totalPaginas > 1 && (
          <div className="flex justify-center gap-2 p-4 border-t border-[#2a2a3e]">
            <button 
              onClick={() => setPaginaActual(Math.max(1, paginaActual - 1))}
              disabled={paginaActual === 1}
              className="px-4 py-2 rounded-lg bg-slate-800 text-white text-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-700 transition-all"
            >
              Anterior
            </button>
            <span className="px-4 py-2 text-slate-400 text-sm">
              Página {paginaActual} de {totalPaginas}
            </span>
            <button 
              onClick={() => setPaginaActual(Math.min(totalPaginas, paginaActual + 1))}
              disabled={paginaActual === totalPaginas}
              className="px-4 py-2 rounded-lg bg-slate-800 text-white text-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-700 transition-all"
            >
              Siguiente
            </button>
          </div>
        )}
      </div>

      <div className="mt-6 grid grid-cols-3 gap-4">
        <div className="bg-gradient-to-br from-blue-500/10 to-blue-600/5 rounded-xl p-4 text-center border border-blue-500/30">
          <div className="text-2xl font-bold text-blue-400">{movimientosFiltrados.length}</div>
          <div className="text-xs text-slate-500">Total Movimientos</div>
        </div>
        <div className="bg-gradient-to-br from-green-500/10 to-green-600/5 rounded-xl p-4 text-center border border-green-500/30">
          <div className="text-2xl font-bold text-green-400">{movimientosFiltrados.filter(m => m.tipo === "entrada").reduce((s, m) => s + m.cantidad, 0)}</div>
          <div className="text-xs text-slate-500">Unidades Entradas</div>
        </div>
        <div className="bg-gradient-to-br from-red-500/10 to-red-600/5 rounded-xl p-4 text-center border border-red-500/30">
          <div className="text-2xl font-bold text-red-400">{movimientosFiltrados.filter(m => m.tipo === "salida").reduce((s, m) => s + m.cantidad, 0)}</div>
          <div className="text-xs text-slate-500">Unidades Salidas</div>
        </div>
      </div>
    </div>
  );
}

// ==================== COMPONENTE PRINCIPAL ====================

export default function Home() {
  const [usuario, setUsuario] = useState<Usuario | null>(null);
  const [productos, setProductos] = useState<Producto[]>([]);
  const [movimientos, setMovimientos] = useState<Movimiento[]>([]);
  const [cargando, setCargando] = useState(true);
  const [busqueda, setBusqueda] = useState("");
  const [categoriaFiltro, setCategoriaFiltro] = useState<Categoria | null>(null);
  const [pestanaActiva, setPestanaActiva] = useState<"ingreso" | "salida">("ingreso");
  const [modalAbierto, setModalAbierto] = useState<ModalData["tipo"]>(null);
  const [modalEstadisticas, setModalEstadisticas] = useState(false);
  const [modalReportes, setModalReportes] = useState(false);
  const [modalAdminProductos, setModalAdminProductos] = useState(false);
  const [modalEditarBase, setModalEditarBase] = useState<Producto | null>(null);
  const [modalRestaurarBackup, setModalRestaurarBackup] = useState(false);
  const [filtroMovimientos, setFiltroMovimientos] = useState<TipoReporte>("todos");
  const [vistaActual, setVistaActual] = useState<Vista>("dashboard");
  const [notificacion, setNotificacion] = useState<Notificacion | null>(null);
  const [busquedaAvanzada, setBusquedaAvanzada] = useState("");
  const [filtroUbicacion, setFiltroUbicacion] = useState("");

  const [productoEntrada, setProductoEntrada] = useState("");
  const [cantidadEntrada, setCantidadEntrada] = useState("");
  const [fechaEntrada, setFechaEntrada] = useState(new Date().toISOString().split('T')[0]);
  const [notasEntrada, setNotasEntrada] = useState("");
  const [motivoEntrada, setMotivoEntrada] = useState("Compra");
  const [loadingEntrada, setLoadingEntrada] = useState(false);
  
  const [productoSalida, setProductoSalida] = useState("");
  const [cantidadSalida, setCantidadSalida] = useState("");
  const [notasSalida, setNotasSalida] = useState("");
  const [motivoSalida, setMotivoSalida] = useState("Venta");
  const [loadingSalida, setLoadingSalida] = useState(false);

  const mostrarNotificacion = useCallback((tipo: Notificacion['tipo'], mensaje: string) => {
    setNotificacion({ tipo, mensaje, id: Date.now() });
  }, []);

  const cargarProductos = useCallback(async () => { 
    const { data } = await supabase.from("productos").select("*").order("nombre"); 
    if (data) setProductos(data); 
  }, []);

  const cargarMovimientos = useCallback(async () => { 
    const { data } = await supabase.from("movimientos").select("*").order("created_at", { ascending: false }).limit(500); 
    if (data) setMovimientos(data); 
  }, []);

  // Backup automático
  useEffect(() => {
    const backupInterval = setInterval(() => {
      if (productos.length > 0) {
        const backup = {
          productos,
          movimientos,
          fecha: new Date().toISOString()
        };
        localStorage.setItem(`backup_${new Date().toISOString().slice(0, 19)}`, JSON.stringify(backup));
        
        const backups = Object.keys(localStorage).filter(k => k.startsWith('backup_'));
        if (backups.length > 10) {
          backups.sort().slice(0, -10).forEach(k => localStorage.removeItem(k));
        }
        
        console.log('Backup automático guardado');
      }
    }, 5 * 60 * 1000);
    
    return () => clearInterval(backupInterval);
  }, [productos, movimientos]);

  // Restaurar backup
  const handleRestaurarBackup = useCallback(async (productosBackup: Producto[], movimientosBackup: Movimiento[]) => {
    setCargando(true);
    try {
      await supabase.from("productos").delete().neq("id", 0);
      await supabase.from("movimientos").delete().neq("id", 0);
      
      for (const p of productosBackup) {
        await supabase.from("productos").insert({
          nombre: p.nombre,
          categoria: p.categoria,
          stock: p.stock,
          stock_minimo: p.stock_minimo,
          cantidad_base: p.cantidad_base,
          unidades_por_caja: p.unidades_por_caja || 1,
          ubicacion: p.ubicacion
        });
      }
      
      for (const m of movimientosBackup) {
        await supabase.from("movimientos").insert({
          producto_id: m.producto_id,
          tipo: m.tipo,
          cantidad: m.cantidad,
          motivo: m.motivo,
          notas: m.notas,
          created_at: m.created_at
        });
      }
      
      await cargarProductos();
      await cargarMovimientos();
      mostrarNotificacion('success', "Backup restaurado correctamente");
    } catch (error) {
      mostrarNotificacion('error', "Error al restaurar backup");
    } finally {
      setCargando(false);
    }
  }, [cargarProductos, cargarMovimientos, mostrarNotificacion]);

  useEffect(() => {
    const filtroGuardado = localStorage.getItem('ultimo_filtro_categoria');
    if (filtroGuardado && Object.values(Categoria).includes(filtroGuardado as Categoria)) {
      setCategoriaFiltro(filtroGuardado as Categoria);
    }
  }, []);

  useEffect(() => {
    if (categoriaFiltro) {
      localStorage.setItem('ultimo_filtro_categoria', categoriaFiltro);
    }
  }, [categoriaFiltro]);

  useEffect(() => {
    const guardado = localStorage.getItem("usuario_activo");
    if (guardado) { 
      try { 
        const u = JSON.parse(guardado); 
        if (u && u.nombre && u.id) setUsuario(u); 
      } catch {} 
    }
  }, []);

  useEffect(() => { 
    if (usuario) { 
      (async () => { 
        await cargarProductos(); 
        await cargarMovimientos(); 
        setCargando(false); 
      })(); 
    } 
  }, [usuario, cargarProductos, cargarMovimientos]);

  // Atajos de teclado
  useEffect(() => {
    const handleKeyPress = (e: KeyboardEvent) => {
      if (e.ctrlKey && e.key === 'i') {
        e.preventDefault();
        setVistaActual('inventario');
        mostrarNotificacion('info', 'Cambiado a vista de Inventario (Ctrl+I)');
      } else if (e.ctrlKey && e.key === 'm') {
        e.preventDefault();
        setVistaActual('movimientos');
        mostrarNotificacion('info', 'Cambiado a vista de Movimientos (Ctrl+M)');
      } else if (e.ctrlKey && e.key === 'd') {
        e.preventDefault();
        setVistaActual('dashboard');
        mostrarNotificacion('info', 'Cambiado a Dashboard (Ctrl+D)');
      } else if (e.ctrlKey && e.key === 'n') {
        e.preventDefault();
        setPestanaActiva('ingreso');
        mostrarNotificacion('info', 'Modo: Registrar Ingreso (Ctrl+N)');
      } else if (e.ctrlKey && e.key === 's') {
        e.preventDefault();
        setPestanaActiva('salida');
        mostrarNotificacion('info', 'Modo: Registrar Salida (Ctrl+S)');
      } else if (e.ctrlKey && e.key === 'e') {
        e.preventDefault();
        setModalEstadisticas(true);
        mostrarNotificacion('info', 'Abriendo estadísticas...');
      } else if (e.ctrlKey && e.key === 'r') {
        e.preventDefault();
        setModalReportes(true);
        mostrarNotificacion('info', 'Abriendo reportes...');
      }
    };
    
    window.addEventListener('keydown', handleKeyPress);
    return () => window.removeEventListener('keydown', handleKeyPress);
  }, [mostrarNotificacion]);

  const productosFiltradosAvanzado = useMemo(() => {
    return productos.filter(p => {
      const matchNombre = p.nombre.toLowerCase().includes(busquedaAvanzada.toLowerCase());
      const matchCategoria = !categoriaFiltro || p.categoria === categoriaFiltro;
      const matchUbicacion = !filtroUbicacion || p.ubicacion?.toLowerCase().includes(filtroUbicacion.toLowerCase());
      return matchNombre && matchCategoria && matchUbicacion;
    });
  }, [productos, busquedaAvanzada, categoriaFiltro, filtroUbicacion]);

  const totalProductos = productos.length;
  const stockCritico = productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length;
  const sinStock = productos.filter(p => p.stock === 0).length;
  const stockOptimo = productos.filter(p => p.stock > p.stock_minimo * 1.5).length;
  const totalPedidos = productos.reduce((sum, p) => sum + calcularPedidoSugerido(p.stock, p.cantidad_base), 0);
  const entradasHoy = movimientos.filter(m => m.tipo === "entrada" && new Date(m.created_at).toDateString() === new Date().toDateString()).reduce((s, m) => s + m.cantidad, 0);
  const salidasHoy = movimientos.filter(m => m.tipo === "salida" && new Date(m.created_at).toDateString() === new Date().toDateString()).reduce((s, m) => s + m.cantidad, 0);
  const movimientosFiltradosHome = movimientos.filter(m => filtroMovimientos === "entradas" ? m.tipo === "entrada" : filtroMovimientos === "salidas" ? m.tipo === "salida" : true);

  const registrarEntrada = async () => {
    if (!productoEntrada || !cantidadEntrada) { 
      mostrarNotificacion('error', "Completa todos los campos");
      return; 
    }
    
    const cantidadValid = validateCantidad(cantidadEntrada);
    if (!cantidadValid.valida) {
      mostrarNotificacion('error', cantidadValid.error || "Cantidad inválida");
      return;
    }
    
    const prod = productos.find(p => p.nombre === productoEntrada);
    if (!prod) return;
    
    setLoadingEntrada(true);
    const cantidad = parseInt(cantidadEntrada);
    
    try {
      await supabase.from("productos").update({ stock: prod.stock + cantidad }).eq("id", prod.id);
      await supabase.from("movimientos").insert({ 
        producto_id: prod.id, 
        tipo: "entrada", 
        cantidad, 
        motivo: motivoEntrada, 
        fecha: fechaEntrada, 
        notas: notasEntrada,
        usuario: usuario?.nombre
      });
      
      setProductoEntrada(""); 
      setCantidadEntrada(""); 
      setNotasEntrada("");
      setMotivoEntrada("Compra");
      setFechaEntrada(new Date().toISOString().split('T')[0]);
      await cargarProductos(); 
      await cargarMovimientos();
      mostrarNotificacion('success', `✅ Ingreso de ${cantidad} unidades registrado correctamente`);
    } catch (error) {
      mostrarNotificacion('error', "Error al registrar el ingreso");
    } finally {
      setLoadingEntrada(false);
    }
  };

  const registrarSalida = async () => {
    if (!productoSalida || !cantidadSalida) { 
      mostrarNotificacion('error', "Completa todos los campos");
      return; 
    }
    
    const cantidadValid = validateCantidad(cantidadSalida);
    if (!cantidadValid.valida) {
      mostrarNotificacion('error', cantidadValid.error || "Cantidad inválida");
      return;
    }
    
    const prod = productos.find(p => p.nombre === productoSalida);
    if (!prod) return;
    
    const cantidad = parseInt(cantidadSalida);
    if (cantidad > prod.stock) { 
      mostrarNotificacion('error', `Stock insuficiente. Solo hay ${prod.stock} unidades disponibles`);
      return; 
    }
    
    setLoadingSalida(true);
    
    try {
      await supabase.from("productos").update({ stock: prod.stock - cantidad }).eq("id", prod.id);
      await supabase.from("movimientos").insert({ 
        producto_id: prod.id, 
        tipo: "salida", 
        cantidad, 
        motivo: motivoSalida, 
        notas: notasSalida,
        usuario: usuario?.nombre
      });
      
      setProductoSalida(""); 
      setCantidadSalida(""); 
      setNotasSalida("");
      setMotivoSalida("Venta");
      await cargarProductos(); 
      await cargarMovimientos();
      mostrarNotificacion('success', `✅ Salida de ${cantidad} unidades registrada correctamente`);
    } catch (error) {
      mostrarNotificacion('error', "Error al registrar la salida");
    } finally {
      setLoadingSalida(false);
    }
  };

  // ==================== FUNCIÓN DE EXPORTACIÓN A EXCEL MEJORADA ====================
  const descargarExcel = async () => {
  try {
    const wb = new ExcelJS.Workbook();
    const sh = wb.addWorksheet("Inventario");
    
    // Número fijo de columnas que usamos
    const NUM_COLUMNAS = 5;
    const letras = ['A', 'B', 'C', 'D', 'E'];
    
    // Función para aplicar bordes solo a celdas específicas
    const aplicarBordesACelda = (row: any, col: number) => {
      const cell = row.getCell(col);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
      };
    };
    
    // ========== TÍTULO PRINCIPAL ==========
    const tituloRow = sh.addRow(["INFORME DE INVENTARIO"]);
    tituloRow.getCell(1).font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
    tituloRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfbbf24" } };
    tituloRow.height = 30;
    sh.mergeCells(`A${tituloRow.number}:E${tituloRow.number}`);
    // Aplicar bordes solo a la celda combinada (que es la A)
    aplicarBordesACelda(tituloRow, 1);
    
    // ========== FECHA ==========
    const fechaRow = sh.addRow([`Fecha de generación: ${new Date().toLocaleString("es-CO")}`]);
    fechaRow.getCell(1).font = { italic: true, size: 11, color: { argb: "FF666666" } };
    sh.mergeCells(`A${fechaRow.number}:E${fechaRow.number}`);
    aplicarBordesACelda(fechaRow, 1);
    
    // ========== USUARIO ==========
    const usuarioRow = sh.addRow([`Generado por: ${usuario?.nombre || "Bartender"}`]);
    usuarioRow.getCell(1).font = { italic: true, size: 11, color: { argb: "FF666666" } };
    sh.mergeCells(`A${usuarioRow.number}:E${usuarioRow.number}`);
    aplicarBordesACelda(usuarioRow, 1);
    
    sh.addRow([]);
    const filaVacia = sh.lastRow;
    if (filaVacia) {
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(filaVacia, i);
      }
    }
    
    // ========== ENCABEZADOS ==========
    const headers = ["Producto", "Unid./Caja", "Stock Actual", "Pedido (Unid.)", "Cajas a Pedir"];
    const headerRow = sh.addRow(headers);
    for (let i = 1; i <= headers.length; i++) {
      const cell = headerRow.getCell(i);
      cell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1f2937" } };
      aplicarBordesACelda(headerRow, i);
    }
    headerRow.height = 25;
    
    // ========== DATOS DE PRODUCTOS ==========
    productos.forEach((p) => {
      const pedidoUnidades = calcularPedidoSugerido(p.stock, p.cantidad_base);
      const { cajas, resto } = calcularCajasNecesarias(pedidoUnidades, p.unidades_por_caja || 1);
      
      let textoCajas = "";
      if (cajas > 0 && resto > 0) {
        textoCajas = `${cajas} caja${cajas !== 1 ? 's' : ''} + ${resto} und`;
      } else if (cajas > 0) {
        textoCajas = `${cajas} caja${cajas !== 1 ? 's' : ''}`;
      } else if (resto > 0) {
        textoCajas = `${resto} unidades sueltas`;
      } else {
        textoCajas = "0";
      }
      
      const row = sh.addRow([
        p.nombre, 
        p.unidades_por_caja || 1, 
        p.stock, 
        pedidoUnidades, 
        textoCajas
      ]);
      
      // Aplicar bordes a todas las celdas de la fila
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(row, i);
      }
      
      // Resaltar productos con pedido sugerido
      if (pedidoUnidades > 0) {
        const cellPedido = row.getCell(4);
        cellPedido.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfee2e2" } };
        cellPedido.font = { color: { argb: "FFef4444" }, bold: true };
        
        const cellCajas = row.getCell(5);
        cellCajas.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfeedd5" } };
        cellCajas.font = { color: { argb: "FFf97316" }, bold: true };
      }
      
      // Resaltar stock crítico
      if (p.stock <= p.stock_minimo) {
        const cellStock = row.getCell(3);
        cellStock.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfee2e2" } };
        cellStock.font = { color: { argb: "FFef4444" }, bold: true };
      }
    });
    
    // ========== FILA DE TOTALES ==========
    const totalStock = productos.reduce((sum, p) => sum + p.stock, 0);
    const totalPedidosUnidades = productos.reduce((sum, p) => sum + calcularPedidoSugerido(p.stock, p.cantidad_base), 0);
    const totalCajas = productos.reduce((sum, p) => {
      const pedido = calcularPedidoSugerido(p.stock, p.cantidad_base);
      const { cajas } = calcularCajasNecesarias(pedido, p.unidades_por_caja || 1);
      return sum + cajas;
    }, 0);
    
    const totalRow = sh.addRow(["TOTALES", "", totalStock, totalPedidosUnidades, `${totalCajas} cajas`]);
    for (let i = 1; i <= NUM_COLUMNAS; i++) {
      const cell = totalRow.getCell(i);
      cell.font = { bold: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFe5e7eb" } };
      aplicarBordesACelda(totalRow, i);
    }
    
    sh.addRow([]);
    const filaVacia2 = sh.lastRow;
    if (filaVacia2) {
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(filaVacia2, i);
      }
    }
    
    // ========== RESULTADOS CLAVE ==========
    const resumenTitle = sh.addRow(["📊 RESULTADOS CLAVE"]);
    const titleCell = resumenTitle.getCell(1);
    titleCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3b82f6" } };
    sh.mergeCells(`A${resumenTitle.number}:E${resumenTitle.number}`);
    aplicarBordesACelda(resumenTitle, 1);
    
    // Filas de resultados
    const resultados = [
      ["Total de productos en inventario", "", "", "", productos.length],
      ["Unidades totales en stock", "", "", "", totalStock],
      ["Unidades que deben pedirse", "", "", "", totalPedidosUnidades],
      ["Cajas totales a pedir", "", "", "", totalCajas],
      ["Productos con stock crítico", "", "", "", productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length],
      ["Productos agotados (stock 0)", "", "", "", productos.filter(p => p.stock === 0).length]
    ];
    
    resultados.forEach(fila => {
      const row = sh.addRow(fila);
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(row, i);
        if (i === 5) {
          row.getCell(i).font = { bold: true };
        }
      }
    });
    
    sh.addRow([]);
    const filaVacia3 = sh.lastRow;
    if (filaVacia3) {
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(filaVacia3, i);
      }
    }
    
    // ========== PRODUCTOS CON STOCK CRÍTICO ==========
    const criticosTitle = sh.addRow(["⚠️ PRODUCTOS CON STOCK CRÍTICO"]);
    const critTitleCell = criticosTitle.getCell(1);
    critTitleCell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    critTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFef4444" } };
    sh.mergeCells(`A${criticosTitle.number}:E${criticosTitle.number}`);
    aplicarBordesACelda(criticosTitle, 1);
    
    const productosCriticos = productos.filter(p => p.stock <= p.stock_minimo);
    if (productosCriticos.length > 0) {
      const critHeader = sh.addRow(["Producto", "Unid./Caja", "Stock Actual", "Stock Mínimo", "Déficit"]);
      for (let i = 1; i <= 5; i++) {
        const cell = critHeader.getCell(i);
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfee2e2" } };
        aplicarBordesACelda(critHeader, i);
      }
      
      productosCriticos.forEach(p => {
        const deficit = p.stock_minimo - p.stock;
        const deficitCajas = Math.ceil(deficit / (p.unidades_por_caja || 1));
        const row = sh.addRow([p.nombre, p.unidades_por_caja || 1, p.stock, p.stock_minimo, `${deficit > 0 ? deficit : 0} und (${deficitCajas} cajas aprox)`]);
        for (let i = 1; i <= 5; i++) {
          aplicarBordesACelda(row, i);
        }
      });
    } else {
      const noCriticosRow = sh.addRow(["✅ No hay productos con stock crítico"]);
      sh.mergeCells(`A${noCriticosRow.number}:E${noCriticosRow.number}`);
      aplicarBordesACelda(noCriticosRow, 1);
    }
    
    sh.addRow([]);
    const filaVacia4 = sh.lastRow;
    if (filaVacia4) {
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(filaVacia4, i);
      }
    }
    
    // ========== PRODUCTOS QUE NECESITAN PEDIDO ==========
    const pedidoTitle = sh.addRow(["🛒 PRODUCTOS QUE NECESITAN PEDIDO"]);
    const pedTitleCell = pedidoTitle.getCell(1);
    pedTitleCell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    pedTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFf97316" } };
    sh.mergeCells(`A${pedidoTitle.number}:E${pedidoTitle.number}`);
    aplicarBordesACelda(pedidoTitle, 1);
    
    const productosConPedido = productos.filter(p => calcularPedidoSugerido(p.stock, p.cantidad_base) > 0);
    if (productosConPedido.length > 0) {
      const pedHeader = sh.addRow(["Producto", "Unid./Caja", "Stock Actual", "Cantidad Base", "Pedido"]);
      for (let i = 1; i <= 5; i++) {
        const cell = pedHeader.getCell(i);
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfeedd5" } };
        aplicarBordesACelda(pedHeader, i);
      }
      
      productosConPedido.forEach(p => {
        const pedidoUnidades = calcularPedidoSugerido(p.stock, p.cantidad_base);
        const { cajas, resto } = calcularCajasNecesarias(pedidoUnidades, p.unidades_por_caja || 1);
        let textoPedido = "";
        if (cajas > 0 && resto > 0) {
          textoPedido = `${cajas} caja${cajas !== 1 ? 's' : ''} + ${resto} und`;
        } else if (cajas > 0) {
          textoPedido = `${cajas} caja${cajas !== 1 ? 's' : ''}`;
        } else {
          textoPedido = `${resto} und`;
        }
        const row = sh.addRow([p.nombre, p.unidades_por_caja || 1, p.stock, p.cantidad_base, textoPedido]);
        for (let i = 1; i <= 5; i++) {
          aplicarBordesACelda(row, i);
        }
      });
      
      sh.addRow([]);
      const filaAntesTotal = sh.lastRow;
      if (filaAntesTotal) {
        for (let i = 1; i <= NUM_COLUMNAS; i++) {
          aplicarBordesACelda(filaAntesTotal, i);
        }
      }
      
      const totalPedidoRow = sh.addRow(["TOTAL A PEDIR", "", "", `${totalPedidosUnidades} unidades`, `${totalCajas} cajas`]);
      for (let i = 1; i <= 5; i++) {
        const cell = totalPedidoRow.getCell(i);
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfeedd5" } };
        aplicarBordesACelda(totalPedidoRow, i);
      }
    } else {
      const noPedidoRow = sh.addRow(["✅ Todos los productos tienen stock suficiente"]);
      sh.mergeCells(`A${noPedidoRow.number}:E${noPedidoRow.number}`);
      aplicarBordesACelda(noPedidoRow, 1);
    }
    
    sh.addRow([]);
    const filaVacia5 = sh.lastRow;
    if (filaVacia5) {
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(filaVacia5, i);
      }
    }
    
    // ========== GUÍA DE COMPRA RÁPIDA ==========
    const guiaTitle = sh.addRow(["📋 GUÍA DE COMPRA RÁPIDA"]);
    const guiaTitleCell = guiaTitle.getCell(1);
    guiaTitleCell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    guiaTitleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF10b981" } };
    sh.mergeCells(`A${guiaTitle.number}:E${guiaTitle.number}`);
    aplicarBordesACelda(guiaTitle, 1);
    
    if (productosConPedido.length > 0) {
      const guiaHeader = sh.addRow(["Producto", "Stock Actual", "Cantidad Base", "Faltante", "Pedir (Cajas)"]);
      for (let i = 1; i <= 5; i++) {
        const cell = guiaHeader.getCell(i);
        cell.font = { bold: true };
        aplicarBordesACelda(guiaHeader, i);
      }
      
      productosConPedido.slice(0, 20).forEach(p => {
        const faltante = calcularPedidoSugerido(p.stock, p.cantidad_base);
        const { cajas, resto } = calcularCajasNecesarias(faltante, p.unidades_por_caja || 1);
        const textoPedido = cajas > 0 ? `${cajas} caja${cajas !== 1 ? 's' : ''}${resto > 0 ? ` + ${resto} und` : ''}` : `${resto} und`;
        const row = sh.addRow([p.nombre, p.stock, p.cantidad_base, faltante, textoPedido]);
        for (let i = 1; i <= 5; i++) {
          aplicarBordesACelda(row, i);
        }
      });
    }
    
    sh.addRow([]);
    const filaVacia6 = sh.lastRow;
    if (filaVacia6) {
      for (let i = 1; i <= NUM_COLUMNAS; i++) {
        aplicarBordesACelda(filaVacia6, i);
      }
    }
    
    // ========== FOOTER ==========
    const footerRow = sh.addRow([`Reporte generado por Sistema de Inventario - ${new Date().getFullYear()}`]);
    footerRow.getCell(1).font = { italic: true, size: 10, color: { argb: "FF888888" } };
    sh.mergeCells(`A${footerRow.number}:E${footerRow.number}`);
    aplicarBordesACelda(footerRow, 1);
    
    // ========== CONFIGURACIÓN DE COLUMNAS ==========
    sh.columns = [
      { width: 35 },  // Producto
      { width: 15 },  // Unid./Caja
      { width: 15 },  // Stock Actual
      { width: 18 },  // Pedido Unidades
      { width: 30 }   // Cajas a Pedir
    ];
    
    // Generar y descargar
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `inventario_${new Date().toLocaleDateString("es-CO").replace(/\//g, "-")}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
    
    mostrarNotificacion('success', "✅ Reporte Excel generado exitosamente");
  } catch (error) {
    console.error("Error al exportar:", error);
    mostrarNotificacion('error', "❌ Error al generar el reporte Excel");
  }
};

  const now = new Date();
  const hora = now.toLocaleTimeString("es-CO", { hour: "2-digit", minute: "2-digit" });
  const fecha = now.toLocaleDateString("es-CO", { day: "numeric", month: "long", year: "numeric" });
  const handleCardClick = useCallback((tipo: ModalData["tipo"]) => setModalAbierto(tipo), []);
  const handleLogout = useCallback(() => { 
    localStorage.removeItem("usuario_activo"); 
    setUsuario(null); 
    mostrarNotificacion('info', "Sesión cerrada correctamente");
  }, [mostrarNotificacion]);

  // Datos para gráficos del dashboard
  const stockPorCategoria = CATEGORIAS.map(cat => ({
    name: cat,
    stock: productos.filter(p => p.categoria === cat).reduce((sum, p) => sum + p.stock, 0),
  }));

  const estadoProductos = [
    { name: "Óptimo", value: stockOptimo, color: "#22c55e" },
    { name: "Crítico", value: stockCritico, color: "#f97316" },
    { name: "Agotado", value: sinStock, color: "#ef4444" },
  ];

  if (!usuario) return <LoginScreen onLogin={setUsuario} />;
  if (cargando) return (
    <div className="min-h-screen bg-[#0a0d16] flex items-center justify-center">
      <div className="text-center">
        <div className="animate-spin rounded-full h-12 w-12 border-4 border-amber-500 border-t-transparent mx-auto mb-4"></div>
        <p className="text-slate-400">Cargando sistema...</p>
      </div>
    </div>
  );

  const renderContenido = () => {
    switch (vistaActual) {
      case "inventario":
        return <VistaInventario productos={productos} onEditarCantidadBase={(p) => setModalEditarBase(p)} />;
      case "movimientos":
        return <VistaMovimientos movimientos={movimientos} productos={productos} />;
      default:
        return (
          <div className="p-6">
            {/* Tarjetas de métricas */}
            <div className="grid grid-cols-5 gap-5 mb-8">
              {[
                { label: "Total Productos", value: totalProductos, sub: "Productos registrados", color: "#3b82f6", bg: "#1e3a5f", icon: "📦", border: "#3b82f6", tipo: "total" as const },
                { label: "Stock Crítico", value: stockCritico, sub: "Requieren reposición", color: "#f97316", bg: "#3a1e0a", icon: "⚠️", border: "#f97316", tipo: "critico" as const },
                { label: "Agotados", value: sinStock, sub: "Sin stock disponible", color: "#ef4444", bg: "#3a0a0a", icon: "🚫", border: "#ef4444", tipo: "agotados" as const },
                { label: "Stock Óptimo", value: stockOptimo, sub: "Nivel adecuado", color: "#22c55e", bg: "#0a3a1e", icon: "✅", border: "#22c55e", tipo: "optimo" as const },
                { label: "Pedido Sugerido", value: totalPedidos, sub: "Unidades a pedir", color: "#f97316", bg: "#3a1e0a", icon: "📦", border: "#f97316", tipo: "pedidos" as const },
              ].map((card) => (
                <div key={card.label} onClick={() => handleCardClick(card.tipo)} className="bg-[#0f1117] border rounded-xl p-5 cursor-pointer hover:scale-105 transition-all duration-200 shadow-lg hover:shadow-xl" style={{ borderColor: `${card.border}40` }}>
                  <div className="flex items-center justify-between mb-3">
                    <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">{card.label}</span>
                    <div className="w-10 h-10 rounded-xl flex items-center justify-center text-xl" style={{ background: `${card.color}20`, color: card.color }}>{card.icon}</div>
                  </div>
                  <div className="text-3xl font-bold text-white mb-1">{card.value}</div>
                  <div className="text-xs text-slate-500">{card.sub}</div>
                </div>
              ))}
            </div>

            {/* Gráficos del dashboard */}
            <div className="grid grid-cols-2 gap-6 mb-8">
              <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4">
                <h3 className="text-sm font-semibold text-white mb-4 flex items-center gap-2">📊 Stock por Categoría</h3>
                <div style={{ height: 300 }}>
                  <ResponsiveContainer>
                    <BarChart data={stockPorCategoria}>
                      <CartesianGrid stroke="#2a2a3e" strokeDasharray="3 3" />
                      <XAxis dataKey="name" stroke="#94a3b8" fontSize={10} angle={-45} textAnchor="end" height={80} />
                      <YAxis stroke="#94a3b8" fontSize={12} />
                      <Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} />
                      <Bar dataKey="stock" fill="#22c55e" radius={[8,8,0,0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              
              <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4">
                <h3 className="text-sm font-semibold text-white mb-4 flex items-center gap-2">🥧 Estado del Inventario</h3>
                <div style={{ height: 300 }}>
                  <ResponsiveContainer>
                    <PieChart>
                      <Pie 
                        data={estadoProductos} 
                        cx="50%" 
                        cy="50%" 
                        labelLine={true} 
                        label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`} 
                        outerRadius={120} 
                        dataKey="value"
                      >
                        {estadoProductos.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={entry.color} />
                        ))}
                      </Pie>
                      <Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} />
                      <Legend />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Búsqueda y filtros */}
            <div className="flex gap-3 mb-6 flex-wrap items-center">
              <div className="relative flex-1 max-w-md">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">🔍</span>
                <input 
                  type="text" 
                  placeholder="Buscar producto por nombre..." 
                  value={busqueda} 
                  onChange={(e) => setBusqueda(e.target.value)} 
                  className="pl-9 pr-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-full focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20 transition-all" 
                />
              </div>
              <button onClick={() => setCategoriaFiltro(null)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all ${!categoriaFiltro ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>Todas las categorías</button>
              {CATEGORIAS.map(cat => (
                <button key={cat} onClick={() => setCategoriaFiltro(cat)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all flex items-center gap-2 ${categoriaFiltro === cat ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>
                  <span>{CATEGORIA_EMOJI[cat]}</span>{cat}
                </button>
              ))}
            </div>

            {/* Tabla de productos */}
            <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
              <div className="px-5 py-4 flex items-center justify-between border-b border-[#2a2a3e] bg-[#0a0d16]">
                <span className="text-base font-semibold text-white">📋 Productos del inventario</span>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-sm">🔍</span>
                  <input 
                    placeholder="Filtrar productos..." 
                    value={busqueda} 
                    onChange={(e) => setBusqueda(e.target.value)} 
                    className="pl-9 pr-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-64 focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20" 
                  />
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[800px]">
                  <thead className="bg-[#0a0d16] border-b border-[#2a2a3e]">
                    <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wider">
                      <th className="p-4">Producto</th>
                      <th className="p-4 text-center">Stock</th>
                      <th className="p-4 text-center">Base</th>
                      <th className="p-4 text-center">Pedido</th>
                      <th className="p-4">Estado</th>
                      <th className="p-4">Acción</th>
                    </tr>
                  </thead>
                  <tbody>
                    {productosFiltradosAvanzado.slice(0, 10).map((p, i) => {
                      const est = estadoInfo(p);
                      const pedido = calcularPedidoSugerido(p.stock, p.cantidad_base);
                      const { cajas, resto } = calcularCajasNecesarias(pedido, p.unidades_por_caja || 1);
                      const textoPedido = cajas > 0 ? `${cajas} caja${cajas !== 1 ? 's' : ''}${resto > 0 ? ` + ${resto}` : ''}` : `${resto} und`;
                      return (
                        <tr key={p.id} className={`border-b border-[#1a1a2a] hover:bg-[#1a1a2a] transition-all ${i % 2 === 0 ? 'bg-transparent' : 'bg-[#0a0d16]/50'}`}>
                          <td className="p-4">
                            <div>
                              <span className="text-white font-medium text-base">{p.nombre}</span>
                              <span className="text-xs ml-2 px-2 py-1 rounded-full" style={{ background: CATEGORIA_BG[p.categoria], color: CATEGORIA_COLOR[p.categoria], border: `1px solid ${CATEGORIA_BORDER[p.categoria]}` }}>
                                {CATEGORIA_EMOJI[p.categoria]} {p.categoria}
                              </span>
                            </div>
                          </td>
                          <td className="p-4 text-center"><span className="text-2xl font-bold" style={{ color: est.color }}>{p.stock}</span></td>
                          <td className="p-4 text-center"><span className="text-lg font-semibold text-blue-400">{p.cantidad_base}</span></td>
                          <td className="p-4 text-center">{pedido > 0 ? <span className="text-xl font-bold text-orange-400">{textoPedido}</span> : <span className="text-sm text-green-500">✓</span>}</td>
                          <td className="p-4"><span className="text-xs px-3 py-1.5 rounded-full font-semibold" style={{ background: est.bg, color: est.color, border: `1px solid ${est.color}40` }}><span className="inline-block w-2 h-2 rounded-full mr-1.5" style={{ background: est.dot }} />{est.label}</span></td>
                          <td className="p-4">
                            <div className="flex gap-2">
                              <button onClick={() => { setPestanaActiva("salida"); setProductoSalida(p.nombre); }} className="px-3 py-1.5 rounded-lg bg-red-500/10 border border-red-500/30 text-red-400 text-xs font-medium hover:bg-red-500/20 transition-all">Salida</button>
                              <button onClick={() => setModalEditarBase(p)} className="px-3 py-1.5 rounded-lg bg-amber-500/10 border border-amber-500/30 text-amber-400 text-xs font-medium hover:bg-amber-500/20 transition-all">Editar</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {productosFiltradosAvanzado.length === 0 && (
                      <tr>
                        <td colSpan={6} className="text-center py-12 text-slate-500">No se encontraron productos</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Últimos movimientos */}
            <div className="mt-8">
              <div className="flex justify-between items-center mb-4">
                <h3 className="text-base font-semibold text-white flex items-center gap-2">📋 Últimos movimientos registrados</h3>
                <div className="flex gap-2">
                  <button onClick={() => setFiltroMovimientos("todos")} className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-all ${filtroMovimientos === "todos" ? 'bg-amber-600 text-white' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>Todos</button>
                  <button onClick={() => setFiltroMovimientos("entradas")} className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-all flex items-center gap-1 ${filtroMovimientos === "entradas" ? 'bg-green-600 text-white' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}><span>↑</span> Entradas</button>
                  <button onClick={() => setFiltroMovimientos("salidas")} className={`px-4 py-1.5 rounded-lg text-sm font-medium transition-all flex items-center gap-1 ${filtroMovimientos === "salidas" ? 'bg-red-600 text-white' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}><span>↓</span> Salidas</button>
                </div>
              </div>
              <div className="flex gap-3 flex-wrap">
                {movimientosFiltradosHome.slice(0, 12).map(m => {
                  const prod = productos.find(p => p.id === m.producto_id);
                  const isEntrada = m.tipo === "entrada";
                  return (
                    <div key={m.id} className="bg-[#0f1117] border rounded-xl px-4 py-2.5 text-sm flex items-center gap-3 shadow-md animate-in slide-in-from-left duration-300" style={{ borderColor: isEntrada ? '#166534' : '#991b1b' }}>
                      <span className="text-xl font-bold" style={{ color: isEntrada ? '#22c55e' : '#ef4444' }}>{isEntrada ? "↑" : "↓"}</span>
                      <span className="font-bold text-lg" style={{ color: isEntrada ? '#22c55e' : '#ef4444' }}>{m.cantidad}</span>
                      <span className="text-white font-medium">{prod?.nombre || "?"}</span>
                      <span className="text-slate-500 text-xs">{new Date(m.created_at).toLocaleDateString("es-CO")}</span>
                      {m.motivo && m.motivo !== "Compra" && <span className="bg-slate-800 px-2 py-1 rounded-lg text-xs text-slate-400">{m.motivo}</span>}
                    </div>
                  );
                })}
                {movimientosFiltradosHome.length === 0 && <div className="text-center py-8 text-slate-500 w-full">No hay movimientos registrados</div>}
              </div>
            </div>
          </div>
        );
    }
  };

  return (
    <div className="min-h-screen bg-[#0a0d16]">
      {notificacion && (
        <NotificationToast notificacion={notificacion} onClose={() => setNotificacion(null)} />
      )}
      
      <header className="bg-[#0f1117] border-b border-[#2a2a3e] px-6 py-4 flex items-center justify-between sticky top-0 z-50 shadow-lg">
        <div className="flex items-center gap-4">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-r from-amber-500 to-amber-600 flex items-center justify-center text-xl shadow-md">🍺</div>
          <div>
            <h1 className="text-xl font-bold bg-gradient-to-r from-amber-400 to-amber-600 bg-clip-text text-transparent">INVENTARIO</h1>
            <p className="text-xs text-slate-400">Sistema profesional de control de inventario</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex gap-6 px-4 py-2 bg-[#0a0d16] rounded-xl border border-[#2a2a3e]">
            <span className="text-sm text-slate-400">↑ Entradas hoy: <strong className="text-green-500 text-lg ml-1">{entradasHoy}</strong></span>
            <span className="text-sm text-slate-400">↓ Salidas hoy: <strong className="text-red-500 text-lg ml-1">{salidasHoy}</strong></span>
          </div>
          <button onClick={() => setModalEstadisticas(true)} className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-sm font-medium transition-all flex items-center gap-2 shadow-md">📊 Estadísticas</button>
          <button onClick={() => setModalReportes(true)} className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-sm font-medium transition-all flex items-center gap-2 shadow-md">📋 Reportes</button>
          <button onClick={descargarExcel} className="px-4 py-2 rounded-xl bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white text-sm font-medium transition-all flex items-center gap-2 shadow-md">📥 Exportar Excel</button>
          <button onClick={() => setModalRestaurarBackup(true)} className="px-4 py-2 rounded-xl bg-purple-500/20 border border-purple-500/30 text-purple-400 text-sm font-medium hover:bg-purple-500/30 transition-all">💾 Backup</button>
          <div className="text-right border-l border-[#2a2a3e] pl-4">
            <div className="text-base font-semibold text-white">{hora}</div>
            <div className="text-xs text-slate-500">{fecha}</div>
          </div>
          <button onClick={handleLogout} className="px-4 py-2 rounded-xl bg-red-500/10 border border-red-500/30 text-red-400 text-sm font-medium hover:bg-red-500/20 transition-all">Salir</button>
        </div>
      </header>

      <div className="flex h-[calc(100vh-73px)]">
        <aside className="w-64 bg-[#0f1117] border-r border-[#2a2a3e] flex flex-col flex-shrink-0">
          <div className="p-5 border-b border-[#2a2a3e]">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-amber-500/20 flex items-center justify-center text-lg"></div>
              <div>
                <div className="text-sm font-semibold text-white">INVENTARIO</div>
                <div className="text-[10px] text-slate-500">Version 2.0</div>
              </div>
            </div>
          </div>
          <nav className="flex-1 p-3 space-y-1">
            <div className="px-3 py-2 text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Principal</div>
            <div onClick={() => setVistaActual("dashboard")} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer ${vistaActual === "dashboard" ? 'bg-gradient-to-r from-amber-500/20 to-amber-500/10 text-amber-400' : 'text-slate-400 hover:bg-slate-800'}`}><span>📊</span> Dashboard <span className="ml-auto text-[10px] text-slate-500">Ctrl+D</span></div>
            <div onClick={() => setVistaActual("inventario")} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer ${vistaActual === "inventario" ? 'bg-gradient-to-r from-amber-500/20 to-amber-500/10 text-amber-400' : 'text-slate-400 hover:bg-slate-800'}`}><span>📦</span> Inventario <span className="ml-auto bg-red-500 text-white text-[10px] rounded-full px-2 py-0.5">{productos.filter(p => p.stock <= p.stock_minimo).length}</span> <span className="text-[10px] text-slate-500">Ctrl+I</span></div>
            <div onClick={() => setVistaActual("movimientos")} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer ${vistaActual === "movimientos" ? 'bg-gradient-to-r from-amber-500/20 to-amber-500/10 text-amber-400' : 'text-slate-400 hover:bg-slate-800'}`}><span>🔄</span> Movimientos <span className="text-[10px] text-slate-500">Ctrl+M</span></div>
            <div className="px-3 pt-4 pb-2 text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Administración</div>
            <div onClick={() => setModalAdminProductos(true)} className="flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer text-slate-400 hover:bg-slate-800">
              <span>🛠️</span> Admin Productos
            </div>
            <div className="px-3 pt-4 pb-2 text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Atajos de teclado</div>
            <div className="px-4 py-1 text-xs text-slate-500">Ctrl+I: Inventario</div>
            <div className="px-4 py-1 text-xs text-slate-500">Ctrl+M: Movimientos</div>
            <div className="px-4 py-1 text-xs text-slate-500">Ctrl+D: Dashboard</div>
            <div className="px-4 py-1 text-xs text-slate-500">Ctrl+N: Nuevo ingreso</div>
            <div className="px-4 py-1 text-xs text-slate-500">Ctrl+S: Nueva salida</div>
            <div className="px-4 py-1 text-xs text-slate-500">Ctrl+E: Estadísticas</div>
            <div className="px-4 py-1 text-xs text-slate-500">Ctrl+R: Reportes</div>
          </nav>
          <div className="p-4 border-t border-[#2a2a3e]">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-full bg-gradient-to-r from-amber-500 to-amber-600 flex items-center justify-center text-white font-semibold text-sm">B</div>
              <div className="flex-1">
                <div className="text-sm font-medium text-white">{usuario?.nombre}</div>
                <div className="text-[10px] text-green-500 flex items-center gap-1"><span className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse"></span> En línea</div>
              </div>
              <button onClick={handleLogout} className="text-slate-500 hover:text-red-400 transition-all text-lg" title="Salir">🚪</button>
            </div>
          </div>
        </aside>

        <div className="flex-1 overflow-auto">
          {renderContenido()}
        </div>

        {vistaActual === "dashboard" && (
          <aside className="w-96 bg-[#0f1117] border-l border-[#2a2a3e] flex flex-col flex-shrink-0 shadow-xl">
            <div className="m-5 p-4 rounded-xl bg-gradient-to-r from-orange-500/10 to-orange-600/5 border border-orange-500/30">
              <div className="flex items-center gap-2 text-sm font-semibold text-orange-400 mb-3"><span>⚠️</span> Alertas de stock crítico</div>
              <div className="space-y-2 max-h-64 overflow-auto">
                {productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).slice(0, 5).map(p => (
                  <div key={p.id} className="flex justify-between items-center py-2 px-2 rounded-lg bg-slate-800/50">
                    <span className="text-sm text-slate-300">{p.nombre}</span>
                    <span className="text-orange-400 font-bold text-base">{p.stock} / {p.stock_minimo} min</span>
                  </div>
                ))}
                {productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length === 0 && <div className="text-center py-4 text-slate-500 text-sm">No hay alertas activas</div>}
              </div>
            </div>

            <div className="flex mx-5 bg-[#0a0d16] rounded-xl p-1 gap-1">
              <button onClick={() => setPestanaActiva("ingreso")} className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${pestanaActiva === "ingreso" ? 'bg-gradient-to-r from-green-600 to-green-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}><span>↑</span> Registrar Ingreso</button>
              <button onClick={() => setPestanaActiva("salida")} className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${pestanaActiva === "salida" ? 'bg-gradient-to-r from-red-600 to-red-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}><span>↓</span> Registrar Salida</button>
            </div>

            <div className="flex-1 overflow-auto p-5">
              {pestanaActiva === "ingreso" ? (
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Producto</label>
                    <select value={productoEntrada} onChange={(e) => setProductoEntrada(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500 focus:ring-2 focus:ring-amber-500/20">
                      <option value="">Seleccionar producto...</option>
                      {productos.map(p => <option key={p.id} value={p.nombre}>{p.nombre} — Stock actual: {p.stock}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Cantidad</label>
                    <input type="number" value={cantidadEntrada} onChange={(e) => setCantidadEntrada(e.target.value)} placeholder="Número de unidades" className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Motivo</label>
                    <select value={motivoEntrada} onChange={(e) => setMotivoEntrada(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm">
                      <option value="Compra">Compra a proveedor</option>
                      <option value="Devolución">Devolución de cliente</option>
                      <option value="Ajuste">Ajuste de inventario</option>
                      <option value="Transferencia">Transferencia</option>
                      <option value="Otro">Otro</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Fecha de ingreso</label>
                    <input type="date" value={fechaEntrada} onChange={(e) => setFechaEntrada(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Notas</label>
                    <textarea value={notasEntrada} onChange={(e) => setNotasEntrada(e.target.value)} rows={3} placeholder="Observaciones adicionales..." className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm resize-none focus:outline-none focus:border-amber-500" />
                  </div>
                  <button onClick={registrarEntrada} disabled={loadingEntrada || !productoEntrada || !cantidadEntrada} className="w-full py-3 rounded-xl bg-gradient-to-r from-green-600 to-green-500 hover:from-green-500 hover:to-green-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2 shadow-lg">
                    {loadingEntrada ? (
                      <>
                        <div className="animate-spin rounded-full h-5 w-5 border-2 border-white border-t-transparent"></div>
                        PROCESANDO...
                      </>
                    ) : "↑ REGISTRAR INGRESO"}
                  </button>
                </div>
              ) : (
                <div className="space-y-4">
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Producto</label>
                    <select value={productoSalida} onChange={(e) => setProductoSalida(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500">
                      <option value="">Seleccionar producto...</option>
                      {productos.map(p => <option key={p.id} value={p.nombre}>{p.nombre} — Stock actual: {p.stock}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Cantidad</label>
                    <input type="number" value={cantidadSalida} onChange={(e) => setCantidadSalida(e.target.value)} placeholder="Número de unidades" className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Motivo</label>
                    <select value={motivoSalida} onChange={(e) => setMotivoSalida(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm">
                      <option value="Venta">Venta</option>
                      <option value="Merma">Merma / Pérdida</option>
                      <option value="Transferencia">Transferencia</option>
                      <option value="Cortesía">Cortesía / Muestra</option>
                      <option value="Ajuste">Ajuste de inventario</option>
                      <option value="Otro">Otro</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-slate-400 mb-2">Notas</label>
                    <textarea value={notasSalida} onChange={(e) => setNotasSalida(e.target.value)} rows={3} placeholder="Observaciones..." className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm resize-none focus:outline-none focus:border-amber-500" />
                  </div>
                  {productoSalida && cantidadSalida && (
                    <div className="p-4 rounded-xl bg-red-500/10 border border-red-500/30">
                      <span className="text-sm text-slate-400">Stock restante: </span>
                      <strong className="text-red-400 text-xl ml-2">{Math.max(0, (productos.find(p => p.nombre === productoSalida)?.stock || 0) - parseInt(cantidadSalida || "0"))}</strong>
                      <span className="text-slate-400 text-sm ml-1">unidades</span>
                    </div>
                  )}
                  <button onClick={registrarSalida} disabled={loadingSalida || !productoSalida || !cantidadSalida} className="w-full py-3 rounded-xl bg-gradient-to-r from-red-600 to-red-500 hover:from-red-500 hover:to-red-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2 shadow-lg">
                    {loadingSalida ? (
                      <>
                        <div className="animate-spin rounded-full h-5 w-5 border-2 border-white border-t-transparent"></div>
                        PROCESANDO...
                      </>
                    ) : "↓ REGISTRAR SALIDA"}
                  </button>
                </div>
              )}
            </div>
          </aside>
        )}
      </div>

      {modalAbierto && <ModalDetalle tipo={modalAbierto} productos={productos} onClose={() => setModalAbierto(null)} />}
      {modalEstadisticas && <ModalEstadisticas productos={productos} movimientos={movimientos} onClose={() => setModalEstadisticas(false)} />}
      {modalReportes && <ModalReportes productos={productos} movimientos={movimientos} onClose={() => setModalReportes(false)} />}
      {modalAdminProductos && (
        <ModalAdminProductos 
          onClose={() => setModalAdminProductos(false)} 
          productos={productos}
          onProductoCreado={async () => {
            await cargarProductos();
            await cargarMovimientos();
          }}
          mostrarNotificacion={mostrarNotificacion}
        />
      )}
      {modalEditarBase && (
        <ModalEditarCantidadBase 
          producto={modalEditarBase}
          onClose={() => setModalEditarBase(null)}
          onActualizado={async () => {
            await cargarProductos();
          }}
          mostrarNotificacion={mostrarNotificacion}
        />
      )}
      {modalRestaurarBackup && (
        <ModalRestaurarBackup 
          onClose={() => setModalRestaurarBackup(false)}
          onRestaurar={handleRestaurarBackup}
          mostrarNotificacion={mostrarNotificacion}
        />
      )}
    </div>
  );
}