"use client";

import { supabase } from "@/lib/supabase";
import { useEffect, useState, useMemo, useCallback } from "react";
import * as ExcelJS from "exceljs";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line
} from 'recharts';

// ==================== TIPOS Y ENUMS ====================

type TipoMovimiento = 'entrada' | 'salida';
type TipoReporte = 'todos' | 'entradas' | 'salidas';
type EstadoStock = 'AGOTADO' | 'CRÍTICO' | 'BAJO' | 'ÓPTIMO';
type Vista = "dashboard" | "inventario" | "movimientos";

type Producto = {
  id: number;
  nombre: string;
  categoria: Categoria;
  stock: number;
  stock_minimo: number;
  cantidad_base: number;
  ubicacion?: string;
  created_at?: string;
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
};

type Usuario = {
  id: number;
  nombre: string;
  rol?: 'admin' | 'usuario';
};

type Notificacion = {
  tipo: 'success' | 'error' | 'warning' | 'info';
  mensaje: string;
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
    <div className="fixed top-20 right-6 z-50">
      <div className={`px-6 py-4 rounded-xl border backdrop-blur-md ${colors[notificacion.tipo]} shadow-2xl flex items-center gap-3 min-w-[300px]`}>
        <span className="text-2xl">{icons[notificacion.tipo]}</span>
        <p className="text-sm font-medium">{notificacion.mensaje}</p>
        <button onClick={onClose} className="ml-auto text-slate-400 hover:text-white">✕</button>
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
    <div className="min-h-screen bg-[#0a0d16] flex items-center justify-center p-6">
      <div className="w-full max-w-md">
        <div className="text-center mb-10">
          <div className="inline-flex items-center justify-center w-20 h-20 rounded-2xl bg-gradient-to-br from-amber-500 to-amber-600 mb-5 shadow-lg">
            <span className="text-4xl">🍺</span>
          </div>
          <h1 className="text-3xl font-bold text-white tracking-tight">INVENTARIO</h1>
          <p className="text-sm text-slate-500 mt-2">Sistema profesional de control de inventario</p>
        </div>
        
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl shadow-xl overflow-hidden">
          <div className="px-8 pt-8 pb-5 border-b border-[#2a2a3e]">
            <h2 className="text-xl font-semibold text-white">Iniciar sesión</h2>
            <p className="text-sm text-slate-500 mt-1">Ingresa tus credenciales para acceder al sistema</p>
          </div>
          
          <form onSubmit={handleLogin} className="px-8 pb-8 pt-6 space-y-6">
            {error && (
              <div className="bg-red-500/10 border border-red-500/30 rounded-xl p-4">
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
                className="w-full px-4 py-3 bg-[#0a0d16] border border-[#2a2a3e] rounded-xl text-white text-base placeholder-slate-600 focus:outline-none focus:border-amber-500 focus:ring-1 focus:ring-amber-500 transition-all"
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
                className="w-full px-4 py-3 bg-[#0a0d16] border border-[#2a2a3e] rounded-xl text-white text-base placeholder-slate-600 focus:outline-none focus:border-amber-500 focus:ring-1 focus:ring-amber-500 transition-all"
              />
            </div>
            
            <button 
              type="submit" 
              disabled={cargando} 
              className="w-full py-3 bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white font-semibold rounded-xl transition-all duration-200 disabled:opacity-50 flex items-center justify-center gap-2 text-base"
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
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6" onClick={onClose}>
      <div className="relative w-full max-w-5xl max-h-[85vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] shadow-2xl" onClick={(e) => e.stopPropagation()}>
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
          <table className="w-full">
            <thead className="border-b border-[#2a2a3e]">
              <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wide">
                <th className="pb-4">Producto</th>
                <th className="pb-4">Categoría</th>
                <th className="pb-4 text-right">Stock Actual</th>
                {tipo === "pedidos" && (
                  <>
                    <th className="pb-4 text-right">Cantidad Base</th>
                    <th className="pb-4 text-right">Pedido Sugerido</th>
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
                return (
                  <tr key={p.id} className="border-b border-[#1a1a2a]">
                    <td className="py-3 text-white font-medium">{p.nombre}</td>
                    <td className="py-3"><span className="text-xs px-3 py-1.5 rounded-full" style={{ background: CATEGORIA_BG[p.categoria], color: CATEGORIA_COLOR[p.categoria], border: `1px solid ${CATEGORIA_BORDER[p.categoria]}` }}>{CATEGORIA_EMOJI[p.categoria]} {p.categoria}</span></td>
                    <td className="py-3 text-right text-xl font-bold" style={{ color }}>{p.stock}</td>
                    {tipo === "pedidos" && (
                      <>
                        <td className="py-3 text-right text-lg font-semibold text-blue-400">{p.cantidad_base}</td>
                        <td className="py-3 text-right text-lg font-bold" style={{ color: "#f97316" }}>{pedido}</td>
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
                  <td colSpan={tipo === "pedidos" ? 6 : 5} className="text-center py-12 text-slate-500">
                    {tipo === "pedidos" ? "No hay productos que necesiten pedido" : "No hay productos en esta categoría"}
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {tipo === "pedidos" && productosFiltrados.length > 0 && (
          <div className="p-6 border-t border-[#2a2a3e] bg-[#0a0d16]">
            <div className="flex justify-between items-center">
              <div>
                <span className="text-sm text-slate-400">Total de unidades a pedir: </span>
                <span className="text-2xl font-bold text-orange-400">{totalPedido}</span>
              </div>
              <button 
                onClick={() => {
                  alert(`Pedido sugerido generado:\nTotal: ${totalPedido} unidades\nProductos: ${productosFiltrados.length}`);
                }}
                className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-orange-600 to-orange-500 hover:from-orange-500 hover:to-orange-600 text-white font-medium transition-all flex items-center gap-2"
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
        fecha: fechaStr, 
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

  const COLORS = ['#fbbf24', '#60a5fa', '#c084fc', '#4ade80', '#f472b6', '#f97316'];
  
  const rotacionInventario = useMemo(() => {
    const salidasUltimoMes = movimientos
      .filter(m => m.tipo === 'salida' && 
        new Date(m.created_at) > new Date(Date.now() - 30 * 24 * 60 * 60 * 1000))
      .reduce((sum, m) => sum + m.cantidad, 0);
    const promedioInventario = productos.reduce((sum, p) => sum + p.stock, 0) / productos.length;
    return promedioInventario > 0 ? (salidasUltimoMes / promedioInventario).toFixed(2) : "0";
  }, [movimientos, productos]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6" onClick={onClose}>
      <div className="relative w-full max-w-6xl max-h-[90vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-6 sticky top-0 bg-[#0f1117] z-10 pb-4 border-b border-[#2a2a3e]">
          <h2 className="text-2xl font-bold text-white">📊 Estadísticas y Análisis</h2>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>
        
        <div className="grid grid-cols-2 gap-4 mb-8">
          <div className="bg-[#0a0d16] rounded-xl p-4 text-center border border-[#2a2a3e]">
            <div className="text-2xl font-bold text-amber-400">{rotacionInventario}</div>
            <div className="text-xs text-slate-500">Rotación de Inventario (30 días)</div>
          </div>
          <div className="bg-[#0a0d16] rounded-xl p-4 text-center border border-[#2a2a3e]">
            <div className="text-2xl font-bold text-blue-400">{movimientos.length}</div>
            <div className="text-xs text-slate-500">Total Movimientos</div>
          </div>
        </div>
        
        <div className="mb-8"><h3 className="text-base font-semibold text-slate-300 mb-4">Stock por Categoría</h3><div style={{ height: 400 }}><ResponsiveContainer><BarChart data={consumoPorCategoria}><CartesianGrid stroke="#2a2a3e" /><XAxis dataKey="name" stroke="#94a3b8" fontSize={12} /><YAxis stroke="#94a3b8" fontSize={12} /><Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} /><Legend /><Bar dataKey="stock" name="Stock Actual" fill="#22c55e" radius={[8,8,0,0]} /></BarChart></ResponsiveContainer></div></div>
        
        <div className="mb-8"><h3 className="text-base font-semibold text-slate-300 mb-4">Movimientos Últimos 7 Días</h3><div style={{ height: 400 }}><ResponsiveContainer><LineChart data={ultimos7Dias}><CartesianGrid stroke="#2a2a3e" /><XAxis dataKey="fecha" stroke="#94a3b8" fontSize={12} /><YAxis stroke="#94a3b8" fontSize={12} /><Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} /><Legend /><Line type="monotone" dataKey="entradas" name="Entradas" stroke="#22c55e" strokeWidth={3} dot={{ r: 6, fill: "#22c55e" }} /><Line type="monotone" dataKey="salidas" name="Salidas" stroke="#ef4444" strokeWidth={3} dot={{ r: 6, fill: "#ef4444" }} /></LineChart></ResponsiveContainer></div></div>
        
        <div className="mb-8"><h3 className="text-base font-semibold text-slate-300 mb-4">Top 10 Productos más Vendidos</h3><div style={{ height: 450 }}><ResponsiveContainer><BarChart data={topProductos} layout="vertical"><CartesianGrid stroke="#2a2a3e" /><XAxis type="number" stroke="#94a3b8" fontSize={12} /><YAxis type="category" dataKey="nombre" stroke="#94a3b8" fontSize={12} width={160} /><Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} /><Legend /><Bar dataKey="salidas" name="Unidades Vendidas" fill="#f97316" radius={[0,8,8,0]} /><Bar dataKey="stock" name="Stock Actual" fill="#22c55e" radius={[0,8,8,0]} /></BarChart></ResponsiveContainer></div></div>
        
        <div className="mb-6"><h3 className="text-base font-semibold text-slate-300 mb-4">Distribución de Stock por Categoría</h3><div style={{ height: 450 }}><ResponsiveContainer><PieChart><Pie data={consumoPorCategoria} cx="50%" cy="50%" labelLine={true} label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`} outerRadius={160} fill="#8884d8" dataKey="stock">{consumoPorCategoria.map((entry, index) => (<Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />))}</Pie><Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} /><Legend /></PieChart></ResponsiveContainer></div></div>
        
        <div className="flex justify-end pt-4"><button onClick={onClose} className="px-6 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium transition-all">Cerrar</button></div>
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

  const movimientosFiltrados = useMemo(() => {
    return movimientos.filter(m => {
      const fechaMov = m.created_at?.split('T')[0];
      const cumpleFecha = fechaMov >= fechaInicio && fechaMov <= fechaFin;
      const cumpleTipo = tipoReporte === "todos" || m.tipo === tipoReporte;
      const cumpleProducto = !productoFiltro || m.producto_id === parseInt(productoFiltro);
      return cumpleFecha && cumpleTipo && cumpleProducto;
    });
  }, [movimientos, fechaInicio, fechaFin, tipoReporte, productoFiltro]);

  const exportarReporte = async () => {
    setExportando(true);
    try {
      const wb = new ExcelJS.Workbook();
      const sh = wb.addWorksheet("Reporte_Movimientos");
      const headerRow = sh.addRow(["Fecha", "Producto", "Categoría", "Tipo", "Cantidad", "Motivo", "Notas"]);
      headerRow.font = { bold: true, size: 12 };
      headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1f2937" } };
      headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
      
      movimientosFiltrados.forEach(m => {
        const prod = productos.find(p => p.id === m.producto_id);
        sh.addRow([
          formatFecha(m.created_at), 
          prod?.nombre || "?", 
          prod?.categoria || "?", 
          m.tipo === "entrada" ? "Ingreso" : "Salida", 
          m.cantidad, 
          m.motivo || "-", 
          m.notas || "-"
        ]);
      });
      
      sh.columns = [{ width: 20 }, { width: 35 }, { width: 18 }, { width: 12 }, { width: 12 }, { width: 25 }, { width: 35 }];
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

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6" onClick={onClose}>
      <div className="relative w-full max-w-4xl max-h-[90vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-6 sticky top-0 bg-[#0f1117] z-10 pb-4 border-b border-[#2a2a3e]">
          <h2 className="text-2xl font-bold text-white">📋 Generar Reporte Avanzado</h2>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>
        
        <div className="grid grid-cols-3 gap-5 mb-6">
          <div><label className="block text-sm font-medium text-slate-400 mb-2">Fecha Inicio</label><input type="date" value={fechaInicio} onChange={(e) => setFechaInicio(e.target.value)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" /></div>
          <div><label className="block text-sm font-medium text-slate-400 mb-2">Fecha Fin</label><input type="date" value={fechaFin} onChange={(e) => setFechaFin(e.target.value)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" /></div>
          <div><label className="block text-sm font-medium text-slate-400 mb-2">Tipo de Movimiento</label><select value={tipoReporte} onChange={(e) => setTipoReporte(e.target.value as TipoReporte)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"><option value="todos">Todos</option><option value="entradas">Solo Entradas</option><option value="salidas">Solo Salidas</option></select></div>
        </div>
        
        <div className="mb-6"><label className="block text-sm font-medium text-slate-400 mb-2">Filtrar por Producto (opcional)</label><select value={productoFiltro} onChange={(e) => setProductoFiltro(e.target.value)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"><option value="">Todos los productos</option>{productos.map(p => <option key={p.id} value={p.id}>{p.nombre}</option>)}</select></div>
        
        <div className="mb-6 p-4 rounded-xl bg-amber-500/10 border border-amber-500/30">
          <p className="text-sm text-amber-300">
            📊 Resumen del reporte: {movimientosFiltrados.length} movimientos encontrados | 
            ↑ Entradas: {movimientosFiltrados.filter(m => m.tipo === "entrada").reduce((s, m) => s + m.cantidad, 0)} unidades | 
            ↓ Salidas: {movimientosFiltrados.filter(m => m.tipo === "salida").reduce((s, m) => s + m.cantidad, 0)} unidades
          </p>
        </div>
        
        <div className="flex gap-4 justify-end pt-4">
          <button onClick={onClose} className="px-6 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium transition-all">Cancelar</button>
          <button onClick={exportarReporte} disabled={exportando} className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white font-medium transition-all flex items-center gap-2">
            {exportando ? (
              <>
                <div className="animate-spin rounded-full h-4 w-4 border-2 border-white border-t-transparent"></div>
                Exportando...
              </>
            ) : (
              <>📥 Exportar a Excel</>
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
  const [modo, setModo] = useState<"agregar" | "eliminar">("agregar");
  const [nombre, setNombre] = useState("");
  const [categoria, setCategoria] = useState<Categoria>(Categoria.Cervezas);
  const [stock, setStock] = useState("");
  const [stockMinimo, setStockMinimo] = useState("");
  const [cantidadBase, setCantidadBase] = useState("");
  const [ubicacion, setUbicacion] = useState("");
  const [cargando, setCargando] = useState(false);
  const [productoAEliminar, setProductoAEliminar] = useState<number | null>(null);
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

    setCargando(true);
    
    try {
      const nuevoProducto = {
        nombre: nombre.trim(),
        categoria: categoria,
        stock: parseInt(stock),
        stock_minimo: parseInt(stockMinimo),
        cantidad_base: parseInt(cantidadBase),
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

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6" onClick={onClose}>
      <div className="relative w-full max-w-2xl max-h-[90vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="sticky top-0 bg-[#0f1117] border-b border-[#2a2a3e] p-6 flex justify-between items-center z-10">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl flex items-center justify-center text-2xl bg-amber-500/20">🛠️</div>
            <div>
              <h2 className="text-xl font-bold text-white">Administrar Productos</h2>
              <p className="text-sm text-slate-500">Agrega o elimina productos del inventario</p>
            </div>
          </div>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>

        {errorDetalle && (
          <div className="mx-6 mt-4 p-3 rounded-xl bg-red-500/10 border border-red-500/30">
            <p className="text-sm text-red-400">{errorDetalle}</p>
          </div>
        )}

        <div className="flex mx-6 mt-6 bg-[#0a0d16] rounded-xl p-1 gap-1">
          <button 
            onClick={() => { setModo("agregar"); setErrorDetalle(""); }} 
            className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${modo === "agregar" ? 'bg-gradient-to-r from-green-600 to-green-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}
          >
            ➕ Agregar Producto
          </button>
          <button 
            onClick={() => { setModo("eliminar"); setErrorDetalle(""); }} 
            className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${modo === "eliminar" ? 'bg-gradient-to-r from-red-600 to-red-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}
          >
            🗑️ Eliminar Producto
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
                  className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" 
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
                  disabled={cargando || !nombre || !stock || !stockMinimo || !cantidadBase} 
                  className="w-full py-3 rounded-xl bg-gradient-to-r from-green-600 to-green-500 hover:from-green-500 hover:to-green-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2"
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
                      {p.nombre} - Stock: {p.stock} und - Base: {p.cantidad_base} und
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

// ==================== MODAL EDITAR CANTIDAD BASE ====================

function ModalEditarCantidadBase({ producto, onClose, onActualizado, mostrarNotificacion }: { producto: Producto; onClose: () => void; onActualizado: () => void; mostrarNotificacion: (tipo: Notificacion['tipo'], mensaje: string) => void }) {
  const [cantidadBase, setCantidadBase] = useState(producto.cantidad_base?.toString() || "");
  const [cargando, setCargando] = useState(false);

  const handleGuardar = async () => {
    const nuevaCantidad = parseInt(cantidadBase);
    if (isNaN(nuevaCantidad) || nuevaCantidad < 0) {
      mostrarNotificacion('error', "Ingrese una cantidad válida");
      return;
    }

    setCargando(true);
    const { error } = await supabase
      .from("productos")
      .update({ cantidad_base: nuevaCantidad })
      .eq("id", producto.id);

    if (error) {
      console.error("Error:", error);
      mostrarNotificacion('error', "Error al actualizar la cantidad base");
    } else {
      mostrarNotificacion('success', `✅ Cantidad base de "${producto.nombre}" actualizada a ${nuevaCantidad}`);
      onActualizado();
      onClose();
    }
    setCargando(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6" onClick={onClose}>
      <div className="relative w-full max-w-md rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <h2 className="text-xl font-bold text-white mb-4">Editar Cantidad Base</h2>
        <p className="text-slate-400 mb-4">Producto: <span className="text-white font-semibold">{producto.nombre}</span></p>
        
        <div className="mb-4">
          <label className="block text-sm font-medium text-slate-400 mb-2">Cantidad Base (Objetivo)</label>
          <input 
            type="number" 
            value={cantidadBase} 
            onChange={(e) => setCantidadBase(e.target.value)} 
            className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-lg text-center focus:outline-none focus:border-amber-500"
            autoFocus
          />
          <p className="text-xs text-slate-500 mt-2">Cantidad que siempre debe haber en el bar</p>
          <p className="text-xs text-amber-500 mt-1">Stock actual: {producto.stock} unidades</p>
          <p className="text-xs text-orange-400 mt-1">Pedido sugerido: {calcularPedidoSugerido(producto.stock, parseInt(cantidadBase) || 0)} unidades</p>
        </div>

        <div className="flex gap-3">
          <button onClick={onClose} className="flex-1 py-3 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium">Cancelar</button>
          <button onClick={handleGuardar} disabled={cargando} className="flex-1 py-3 rounded-xl bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white font-semibold">
            {cargando ? "Guardando..." : "Guardar"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ==================== VISTA DE INVENTARIO ====================

function VistaInventario({ productos, onEditarCantidadBase }: { productos: Producto[]; onEditarCantidadBase: (producto: Producto) => void }) {
  const [busqueda, setBusqueda] = useState("");
  const [categoriaFiltro, setCategoriaFiltro] = useState<Categoria | null>(null);
  const [paginaActual, setPaginaActual] = useState(1);
  const itemsPorPagina = 20;

  const productosFiltrados = useMemo(() => {
    return productos.filter(p => 
      p.nombre.toLowerCase().includes(busqueda.toLowerCase()) && 
      (!categoriaFiltro || p.categoria === categoriaFiltro)
    );
  }, [productos, busqueda, categoriaFiltro]);

  const totalPaginas = Math.ceil(productosFiltrados.length / itemsPorPagina);
  const productosPaginados = productosFiltrados.slice(
    (paginaActual - 1) * itemsPorPagina,
    paginaActual * itemsPorPagina
  );

  useEffect(() => {
    setPaginaActual(1);
  }, [busqueda, categoriaFiltro]);

  const totalPedidos = productos.reduce((sum, p) => sum + calcularPedidoSugerido(p.stock, p.cantidad_base), 0);

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-white mb-2">📦 Inventario General</h2>
        <p className="text-slate-500">Listado completo de todos los productos con cantidades actuales y pedidos sugeridos</p>
      </div>

      <div className="flex gap-3 mb-6 flex-wrap items-center">
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">🔍</span>
          <input type="text" placeholder="Buscar producto..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} className="pl-9 pr-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-72 focus:outline-none focus:border-amber-500 transition-all" />
        </div>
        <button onClick={() => setCategoriaFiltro(null)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all ${!categoriaFiltro ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>Todas</button>
        {CATEGORIAS.map(cat => <button key={cat} onClick={() => setCategoriaFiltro(cat)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all flex items-center gap-2 ${categoriaFiltro === cat ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}><span>{CATEGORIA_EMOJI[cat]}</span>{cat}</button>)}
      </div>

      <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
        <div className="px-5 py-4 border-b border-[#2a2a3e] bg-[#0a0d16] flex justify-between items-center">
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
                <th className="p-4 text-center">Stock Actual</th>
                <th className="p-4 text-center">Cantidad Base</th>
                <th className="p-4 text-center">Pedido Sugerido</th>
                <th className="p-4">Estado</th>
                <th className="p-4">Ubicación</th>
                <th className="p-4">Acciones</th>
              </tr>
            </thead>
            <tbody>
              {productosPaginados.map((p, i) => {
                const est = estadoInfo(p);
                const pedidoSugerido = calcularPedidoSugerido(p.stock, p.cantidad_base);
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
                    <td className="p-4 text-center">
                      <span className="text-xl font-semibold text-blue-400">{p.cantidad_base}</span>
                      <span className="text-sm text-slate-500 ml-1">und</span>
                    </td>
                    <td className="p-4 text-center">
                      {pedidoSugerido > 0 ? (
                        <div className="flex flex-col items-center">
                          <span className="text-2xl font-bold text-orange-400">{pedidoSugerido}</span>
                          <span className="text-xs text-orange-500">unidades</span>
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
                        Editar Base
                      </button>
                    </td>
                  </tr>
                );
              })}
              {productosFiltrados.length === 0 && (
                <tr>
                  <td colSpan={7} className="text-center py-12 text-slate-500">No se encontraron productos</td>
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
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-blue-400">{productos.length}</div>
          <div className="text-xs text-slate-500">Total Productos</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-green-400">{productos.reduce((sum, p) => sum + p.stock, 0)}</div>
          <div className="text-xs text-slate-500">Unidades Totales</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-orange-400">{totalPedidos}</div>
          <div className="text-xs text-slate-500">Unidades a Pedir</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-yellow-400">{productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length}</div>
          <div className="text-xs text-slate-500">Stock Crítico</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
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
        <h2 className="text-2xl font-bold text-white mb-2">🔄 Historial de Movimientos</h2>
        <p className="text-slate-500">Registro completo de todas las entradas y salidas de inventario</p>
      </div>

      <div className="grid grid-cols-4 gap-4 mb-6">
        <div>
          <label className="block text-xs font-medium text-slate-400 mb-1">Tipo de Movimiento</label>
          <select value={filtroTipo} onChange={(e) => setFiltroTipo(e.target.value as TipoReporte)} className="w-full px-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500">
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
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-blue-400">{movimientosFiltrados.length}</div>
          <div className="text-xs text-slate-500">Total Movimientos</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-green-400">{movimientosFiltrados.filter(m => m.tipo === "entrada").reduce((s, m) => s + m.cantidad, 0)}</div>
          <div className="text-xs text-slate-500">Unidades Entradas</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
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
  const [filtroMovimientos, setFiltroMovimientos] = useState<TipoReporte>("todos");
  const [vistaActual, setVistaActual] = useState<Vista>("dashboard");
  const [notificacion, setNotificacion] = useState<Notificacion | null>(null);

  const [productoEntrada, setProductoEntrada] = useState("");
  const [cantidadEntrada, setCantidadEntrada] = useState("");
  const [fechaEntrada, setFechaEntrada] = useState(new Date().toISOString().split('T')[0]);
  const [notasEntrada, setNotasEntrada] = useState("");
  const [loadingEntrada, setLoadingEntrada] = useState(false);
  const [productoSalida, setProductoSalida] = useState("");
  const [cantidadSalida, setCantidadSalida] = useState("");
  const [notasSalida, setNotasSalida] = useState("");
  const [loadingSalida, setLoadingSalida] = useState(false);

  const mostrarNotificacion = useCallback((tipo: Notificacion['tipo'], mensaje: string) => {
    setNotificacion({ tipo, mensaje });
  }, []);

  const cargarProductos = useCallback(async () => { 
    const { data } = await supabase.from("productos").select("*").order("nombre"); 
    if (data) setProductos(data); 
  }, []);

  const cargarMovimientos = useCallback(async () => { 
    const { data } = await supabase.from("movimientos").select("*").order("created_at", { ascending: false }).limit(500); 
    if (data) setMovimientos(data); 
  }, []);

  useEffect(() => {
    const backupInterval = setInterval(() => {
      if (productos.length > 0) {
        const backup = {
          productos,
          movimientos,
          fecha: new Date().toISOString()
        };
        localStorage.setItem('backup_automatico', JSON.stringify(backup));
        console.log('Backup automático guardado');
      }
    }, 5 * 60 * 1000);
    
    return () => clearInterval(backupInterval);
  }, [productos, movimientos]);

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

  useEffect(() => {
    const handleKeyPress = (e: KeyboardEvent) => {
      if (e.ctrlKey && e.key === 'i') {
        e.preventDefault();
        setVistaActual('inventario');
        mostrarNotificacion('info', 'Cambiado a vista de Inventario');
      } else if (e.ctrlKey && e.key === 'm') {
        e.preventDefault();
        setVistaActual('movimientos');
        mostrarNotificacion('info', 'Cambiado a vista de Movimientos');
      } else if (e.ctrlKey && e.key === 'd') {
        e.preventDefault();
        setVistaActual('dashboard');
        mostrarNotificacion('info', 'Cambiado a Dashboard');
      }
    };
    
    window.addEventListener('keydown', handleKeyPress);
    return () => window.removeEventListener('keydown', handleKeyPress);
  }, [mostrarNotificacion]);

  const productosFiltrados = useMemo(() => {
    return productos.filter(p => 
      p.nombre.toLowerCase().includes(busqueda.toLowerCase()) && 
      (!categoriaFiltro || p.categoria === categoriaFiltro)
    );
  }, [productos, busqueda, categoriaFiltro]);

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
        motivo: "Compra", 
        fecha: fechaEntrada, 
        notas: notasEntrada 
      });
      
      setProductoEntrada(""); 
      setCantidadEntrada(""); 
      setNotasEntrada(""); 
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
        motivo: "Venta", 
        notas: notasSalida 
      });
      
      setProductoSalida(""); 
      setCantidadSalida(""); 
      setNotasSalida("");
      await cargarProductos(); 
      await cargarMovimientos();
      mostrarNotificacion('success', `✅ Salida de ${cantidad} unidades registrada correctamente`);
    } catch (error) {
      mostrarNotificacion('error', "Error al registrar la salida");
    } finally {
      setLoadingSalida(false);
    }
  };

  const descargarExcel = async () => {
    try {
      const wb = new ExcelJS.Workbook();
      const sh = wb.addWorksheet("Inventario");
      
      const tituloRow = sh.addRow(["INFORME DE INVENTARIO"]);
      tituloRow.font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
      tituloRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfbbf24" } };
      tituloRow.height = 30;
      sh.mergeCells(`A${tituloRow.number}:D${tituloRow.number}`);
      
      const fechaRow = sh.addRow([`Fecha de generación: ${new Date().toLocaleString("es-CO")}`]);
      fechaRow.font = { italic: true, size: 11, color: { argb: "FF666666" } };
      sh.mergeCells(`A${fechaRow.number}:D${fechaRow.number}`);
      
      const usuarioRow = sh.addRow([`Generado por: ${usuario?.nombre || "Bartender"}`]);
      usuarioRow.font = { italic: true, size: 11, color: { argb: "FF666666" } };
      sh.mergeCells(`A${usuarioRow.number}:D${usuarioRow.number}`);
      
      sh.addRow([]);
      
      const headers = ["Producto", "Cantidad Base", "Stock Actual", "Pedido Sugerido"];
      const headerRow = sh.addRow(headers);
      headerRow.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1f2937" } };
      headerRow.height = 25;
      
      productos.forEach((p) => {
        const pedido = calcularPedidoSugerido(p.stock, p.cantidad_base);
        const row = sh.addRow([p.nombre, p.cantidad_base, p.stock, pedido]);
        
        if (pedido > 0) {
          row.getCell(4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfee2e2" } };
          row.getCell(4).font = { color: { argb: "FFef4444" }, bold: true };
        }
        
        if (p.stock <= p.stock_minimo) {
          row.getCell(3).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfee2e2" } };
          row.getCell(3).font = { color: { argb: "FFef4444" }, bold: true };
        }
      });
      
      const totalStock = productos.reduce((sum, p) => sum + p.stock, 0);
      const totalPedidosExcel = productos.reduce((sum, p) => sum + calcularPedidoSugerido(p.stock, p.cantidad_base), 0);
      
      const totalRow = sh.addRow(["TOTALES", "", totalStock, totalPedidosExcel]);
      totalRow.font = { bold: true };
      totalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFe5e7eb" } };
      
      sh.addRow([]);
      
      const resumenTitle = sh.addRow(["📊 RESULTADOS CLAVE"]);
      resumenTitle.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
      resumenTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3b82f6" } };
      sh.mergeCells(`A${resumenTitle.number}:D${resumenTitle.number}`);
      
      sh.addRow(["Total de productos en inventario", "", "", productos.length]);
      sh.addRow(["Unidades totales en stock", "", "", totalStock]);
      sh.addRow(["Unidades que deben pedirse", "", "", totalPedidosExcel]);
      sh.addRow(["Productos con stock crítico", "", "", productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length]);
      sh.addRow(["Productos agotados (stock 0)", "", "", productos.filter(p => p.stock === 0).length]);
      
      sh.addRow([]);
      
      const criticosTitle = sh.addRow(["⚠️ PRODUCTOS CON STOCK CRÍTICO (Requieren atención inmediata)"]);
      criticosTitle.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      criticosTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFef4444" } };
      sh.mergeCells(`A${criticosTitle.number}:D${criticosTitle.number}`);
      
      const productosCriticos = productos.filter(p => p.stock <= p.stock_minimo);
      if (productosCriticos.length > 0) {
        const critHeader = sh.addRow(["Producto", "Stock Actual", "Stock Mínimo", "Déficit"]);
        critHeader.font = { bold: true };
        critHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfee2e2" } };
        
        productosCriticos.forEach(p => {
          const deficit = p.stock_minimo - p.stock;
          sh.addRow([p.nombre, p.stock, p.stock_minimo, deficit > 0 ? deficit : 0]);
        });
      } else {
        sh.addRow(["✅ No hay productos con stock crítico"]);
        sh.mergeCells(`A${sh.lastRow.number}:D${sh.lastRow.number}`);
      }
      
      sh.addRow([]);
      
      const pedidoTitle = sh.addRow(["🛒 PRODUCTOS QUE NECESITAN PEDIDO"]);
      pedidoTitle.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      pedidoTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFf97316" } };
      sh.mergeCells(`A${pedidoTitle.number}:D${pedidoTitle.number}`);
      
      const productosConPedido = productos.filter(p => calcularPedidoSugerido(p.stock, p.cantidad_base) > 0);
      if (productosConPedido.length > 0) {
        const pedHeader = sh.addRow(["Producto", "Stock Actual", "Cantidad Base", "Cantidad a Pedir"]);
        pedHeader.font = { bold: true };
        pedHeader.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfeedd5" } };
        
        productosConPedido.forEach(p => {
          const pedido = calcularPedidoSugerido(p.stock, p.cantidad_base);
          sh.addRow([p.nombre, p.stock, p.cantidad_base, pedido]);
        });
        
        sh.addRow([]);
        const totalPedidoRow = sh.addRow(["TOTAL A PEDIR", "", "", totalPedidosExcel]);
        totalPedidoRow.font = { bold: true };
        totalPedidoRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFfeedd5" } };
      } else {
        sh.addRow(["✅ Todos los productos tienen stock suficiente"]);
        sh.mergeCells(`A${sh.lastRow.number}:D${sh.lastRow.number}`);
      }
      
      sh.addRow([]);
      
      const categoriaTitle = sh.addRow(["📈 ANÁLISIS POR CATEGORÍA"]);
      categoriaTitle.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      categoriaTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8b5cf6" } };
      sh.mergeCells(`A${categoriaTitle.number}:D${categoriaTitle.number}`);
      
      const catHeader = sh.addRow(["Categoría", "Productos", "Stock Total", "Stock Mínimo Promedio"]);
      catHeader.font = { bold: true };
      
      CATEGORIAS.forEach(cat => {
        const productosCat = productos.filter(p => p.categoria === cat);
        const stockTotal = productosCat.reduce((sum, p) => sum + p.stock, 0);
        const stockMinimoProm = productosCat.reduce((sum, p) => sum + p.stock_minimo, 0) / (productosCat.length || 1);
        
        sh.addRow([
          `${CATEGORIA_EMOJI[cat]} ${cat}`,
          productosCat.length,
          stockTotal,
          Math.round(stockMinimoProm)
        ]);
      });
      
      sh.addRow([]);
      
      const movTitle = sh.addRow(["🔄 ACTIVIDAD RECIENTE (Últimos 7 días)"]);
      movTitle.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      movTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF06b6d4" } };
      sh.mergeCells(`A${movTitle.number}:D${movTitle.number}`);
      
      const movHeader = sh.addRow(["Fecha", "Entradas (unidades)", "Salidas (unidades)", "Movimiento Neto"]);
      movHeader.font = { bold: true };
      
      for (let i = 6; i >= 0; i--) {
        const fecha = new Date();
        fecha.setDate(fecha.getDate() - i);
        fecha.setHours(0, 0, 0, 0);
        const fechaStr = fecha.toISOString().split('T')[0];
        const fechaDisplay = fecha.toLocaleDateString("es-CO");
        
        const movimientosDia = movimientos.filter(m => m.created_at?.split('T')[0] === fechaStr);
        const entradas = movimientosDia.filter(m => m.tipo === "entrada").reduce((sum, m) => sum + m.cantidad, 0);
        const salidas = movimientosDia.filter(m => m.tipo === "salida").reduce((sum, m) => sum + m.cantidad, 0);
        const neto = entradas - salidas;
        
        const row = sh.addRow([fechaDisplay, entradas, salidas, neto]);
        if (neto < 0) {
          row.getCell(4).font = { color: { argb: "FFef4444" }, bold: true };
        } else if (neto > 0) {
          row.getCell(4).font = { color: { argb: "FF22c55e" }, bold: true };
        }
      }
      
      sh.addRow([]);
      
      const topTitle = sh.addRow(["🏆 TOP 5 PRODUCTOS MÁS VENDIDOS"]);
      topTitle.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
      topTitle.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFeab308" } };
      sh.mergeCells(`A${topTitle.number}:D${topTitle.number}`);
      
      const topHeader = sh.addRow(["Producto", "Unidades Vendidas", "Stock Actual", "Categoría"]);
      topHeader.font = { bold: true };
      
      const topProductos = productos.map(p => ({
        nombre: p.nombre,
        salidas: movimientos.filter(m => m.producto_id === p.id && m.tipo === "salida").reduce((sum, m) => sum + m.cantidad, 0),
        stock: p.stock,
        categoria: p.categoria
      })).sort((a, b) => b.salidas - a.salidas).slice(0, 5);
      
      topProductos.forEach(p => {
        sh.addRow([p.nombre, p.salidas, p.stock, p.categoria]);
      });
      
      sh.addRow([]);
      
      const footerRow = sh.addRow([`Reporte generado por Sistema de Inventario - ${new Date().getFullYear()}`]);
      footerRow.font = { italic: true, size: 10, color: { argb: "FF888888" } };
      sh.mergeCells(`A${footerRow.number}:D${footerRow.number}`);
      
      sh.columns = [
        { width: 35 },
        { width: 20 },
        { width: 20 },
        { width: 25 }
      ];
      
      sh.eachRow(row => {
        row.eachCell(cell => {
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" }
          };
        });
      });
      
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

            <div className="flex gap-3 mb-6 flex-wrap items-center">
              <div className="relative">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">🔍</span>
                <input type="text" placeholder="Buscar producto..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} className="pl-9 pr-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-72 focus:outline-none focus:border-amber-500 transition-all" />
              </div>
              <button onClick={() => setCategoriaFiltro(null)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all ${!categoriaFiltro ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>Todas las categorías</button>
              {CATEGORIAS.map(cat => <button key={cat} onClick={() => setCategoriaFiltro(cat)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all flex items-center gap-2 ${categoriaFiltro === cat ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}><span>{CATEGORIA_EMOJI[cat]}</span>{cat}</button>)}
            </div>

            <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
              <div className="px-5 py-4 flex items-center justify-between border-b border-[#2a2a3e] bg-[#0a0d16]">
                <span className="text-base font-semibold text-white">📋 Productos del inventario</span>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-sm">🔍</span>
                  <input placeholder="Filtrar productos..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} className="pl-9 pr-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-64 focus:outline-none focus:border-amber-500" />
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
                    {productosFiltrados.slice(0, 10).map((p, i) => {
                      const est = estadoInfo(p);
                      const pedido = calcularPedidoSugerido(p.stock, p.cantidad_base);
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
                          <td className="p-4 text-center">{pedido > 0 ? <span className="text-xl font-bold text-orange-400">{pedido}</span> : <span className="text-sm text-green-500">✓</span>}</td>
                          <td className="p-4"><span className="text-xs px-3 py-1.5 rounded-full font-semibold" style={{ background: est.bg, color: est.color, border: `1px solid ${est.color}40` }}><span className="inline-block w-2 h-2 rounded-full mr-1.5" style={{ background: est.dot }} />{est.label}</span></td>
                          <td className="p-4">
                            <div className="flex gap-2">
                              <button onClick={() => { setPestanaActiva("salida"); setProductoSalida(p.nombre); }} className="px-3 py-1.5 rounded-lg bg-red-500/10 border border-red-500/30 text-red-400 text-xs font-medium hover:bg-red-500/20 transition-all">Salida</button>
                              <button onClick={() => setModalEditarBase(p)} className="px-3 py-1.5 rounded-lg bg-amber-500/10 border border-amber-500/30 text-amber-400 text-xs font-medium hover:bg-amber-500/20 transition-all">Editar Base</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                    {productosFiltrados.length === 0 && (
                      <tr>
                        <td colSpan={6} className="text-center py-12 text-slate-500">No se encontraron productos</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

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
                    <div key={m.id} className="bg-[#0f1117] border rounded-xl px-4 py-2.5 text-sm flex items-center gap-3 shadow-md" style={{ borderColor: isEntrada ? '#166534' : '#991b1b' }}>
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
            <h1 className="text-xl font-bold text-white">INVENTARIO</h1>
            <p className="text-xs text-slate-200">Sistema profesional de control de inventario</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <div className="flex gap-6 px-4 py-2 bg-[#0a0d16] rounded-xl border border-[#2a2a3e]">
            <span className="text-sm text-slate-400">↑ Entradas hoy: <strong className="text-green-500 text-lg ml-1">{entradasHoy}</strong></span>
            <span className="text-sm text-slate-400">↓ Salidas hoy: <strong className="text-red-500 text-lg ml-1">{salidasHoy}</strong></span>
          </div>
          <button onClick={() => setModalEstadisticas(true)} className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-sm font-medium transition-all flex items-center gap-2">📊 Estadísticas</button>
          <button onClick={() => setModalReportes(true)} className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-white text-sm font-medium transition-all flex items-center gap-2">📋 Reportes</button>
          <button onClick={descargarExcel} className="px-4 py-2 rounded-xl bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white text-sm font-medium transition-all flex items-center gap-2">📥 Exportar Excel</button>
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
              <div className="w-9 h-9 rounded-xl bg-amber-500/20 flex items-center justify-center text-lg">🍺</div>
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
            <div onClick={() => setModalAdminProductos(true)} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer text-slate-400 hover:bg-slate-800`}>
              <span>🛠️</span> Admin Productos
            </div>
            <div className="px-3 pt-4 pb-2 text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Análisis</div>
            <div onClick={() => setModalEstadisticas(true)} className="flex items-center gap-3 px-4 py-2.5 rounded-xl text-slate-400 hover:bg-slate-800 transition-all cursor-pointer text-sm"><span>📈</span> Estadísticas</div>
            <div onClick={() => setModalReportes(true)} className="flex items-center gap-3 px-4 py-2.5 rounded-xl text-slate-400 hover:bg-slate-800 transition-all cursor-pointer text-sm"><span>📄</span> Reportes</div>
            <div onClick={descargarExcel} className="flex items-center gap-3 px-4 py-2.5 rounded-xl text-slate-400 hover:bg-slate-800 transition-all cursor-pointer text-sm"><span>📥</span> Exportar Excel</div>
          </nav>
          <div className="p-4 border-t border-[#2a2a3e]">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-full bg-gradient-to-r from-amber-500 to-amber-600 flex items-center justify-center text-white font-semibold text-sm">B</div>
              <div className="flex-1">
                <div className="text-sm font-medium text-white">Bartender</div>
                <div className="text-[10px] text-green-500 flex items-center gap-1"><span className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse"></span> En línea</div>
              </div>
              <button onClick={handleLogout} className="text-slate-500 hover:text-red-400 transition-all text-lg">🚪</button>
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
                    <span className="text-orange-400 font-bold text-base">{p.stock} und</span>
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
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Producto</label><select value={productoEntrada} onChange={(e) => setProductoEntrada(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"><option value="">Seleccionar producto...</option>{productos.map(p => <option key={p.id} value={p.nombre}>{p.nombre} — Stock actual: {p.stock}</option>)}</select></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Cantidad</label><input type="number" value={cantidadEntrada} onChange={(e) => setCantidadEntrada(e.target.value)} placeholder="Número de unidades" className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" /></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Fecha de ingreso</label><input type="date" value={fechaEntrada} onChange={(e) => setFechaEntrada(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" /></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Notas</label><textarea value={notasEntrada} onChange={(e) => setNotasEntrada(e.target.value)} rows={3} placeholder="Observaciones adicionales..." className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm resize-none focus:outline-none focus:border-amber-500" /></div>
                  <button onClick={registrarEntrada} disabled={loadingEntrada || !productoEntrada || !cantidadEntrada} className="w-full py-3 rounded-xl bg-gradient-to-r from-green-600 to-green-500 hover:from-green-500 hover:to-green-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2">
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
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Producto</label><select value={productoSalida} onChange={(e) => setProductoSalida(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"><option value="">Seleccionar producto...</option>{productos.map(p => <option key={p.id} value={p.nombre}>{p.nombre} — Stock actual: {p.stock}</option>)}</select></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Cantidad</label><input type="number" value={cantidadSalida} onChange={(e) => setCantidadSalida(e.target.value)} placeholder="Número de unidades" className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" /></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Notas</label><textarea value={notasSalida} onChange={(e) => setNotasSalida(e.target.value)} rows={3} placeholder="Observaciones..." className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm resize-none focus:outline-none focus:border-amber-500" /></div>
                  {productoSalida && cantidadSalida && (
                    <div className="p-4 rounded-xl bg-red-500/10 border border-red-500/30">
                      <span className="text-sm text-slate-400">Stock restante: </span>
                      <strong className="text-red-400 text-xl ml-2">{Math.max(0, (productos.find(p => p.nombre === productoSalida)?.stock || 0) - parseInt(cantidadSalida || "0"))}</strong>
                      <span className="text-slate-400 text-sm ml-1">unidades</span>
                    </div>
                  )}
                  <button onClick={registrarSalida} disabled={loadingSalida || !productoSalida || !cantidadSalida} className="w-full py-3 rounded-xl bg-gradient-to-r from-red-600 to-red-500 hover:from-red-500 hover:to-red-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2">
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
    </div>
  );
}