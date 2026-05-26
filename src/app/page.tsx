"use client";

import { supabase } from "@/lib/supabase";
import { useEffect, useState } from "react";
import * as ExcelJS from "exceljs";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line
} from 'recharts';

type Producto = {
  id: number;
  nombre: string;
  categoria: string;
  stock: number;
  stock_minimo: number;
  ubicacion?: string;
  sku?: string;
  uso_promedio?: number;
  proveedor?: string;
  ultimo_movimiento?: string;
};

type Movimiento = {
  id: number;
  producto_id: number;
  tipo: string;
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
};

type Vista = "dashboard" | "inventario" | "movimientos";

const CATEGORIAS = ["Cervezas", "Soda & Ginger", "Soda Hatsu", "Gaseosas", "Vinos", "Agua"];

const CATEGORIA_COLOR: Record<string, string> = {
  Cervezas: "#fbbf24",
  "Soda & Ginger": "#60a5fa",
  "Soda Hatsu": "#c084fc",
  Gaseosas: "#4ade80",
  Vinos: "#f472b6",
  Agua: "#60a5fa",
};

const CATEGORIA_BG: Record<string, string> = {
  Cervezas: "#1a1200",
  "Soda & Ginger": "#0a1a2a",
  "Soda Hatsu": "#1a0a1a",
  Gaseosas: "#0a1a0a",
  Vinos: "#1a0a10",
  Agua: "#0a1a2a",
};

const CATEGORIA_BORDER: Record<string, string> = {
  Cervezas: "#92400e",
  "Soda & Ginger": "#1e40af",
  "Soda Hatsu": "#6d28d9",
  Gaseosas: "#166534",
  Vinos: "#9d174d",
  Agua: "#1e40af",
};

const CATEGORIA_EMOJI: Record<string, string> = {
  Cervezas: "🍺",
  "Soda & Ginger": "🫚",
  "Soda Hatsu": "💜",
  Gaseosas: "🥤",
  Vinos: "🍷",
  Agua: "💧",
};

function estadoInfo(p: Producto): { label: string; color: string; bg: string; dot: string; textColor: string } {
  if (p.stock === 0) return { label: "AGOTADO", color: "#ef4444", bg: "#2a0a0a", dot: "#ef4444", textColor: "#fca5a5" };
  if (p.stock <= p.stock_minimo) return { label: "CRÍTICO", color: "#f97316", bg: "#2a1a0a", dot: "#f97316", textColor: "#fdba74" };
  if (p.stock <= p.stock_minimo * 1.5) return { label: "BAJO", color: "#eab308", bg: "#1a1200", dot: "#eab308", textColor: "#fde047" };
  return { label: "ÓPTIMO", color: "#22c55e", bg: "#0a2a14", dot: "#22c55e", textColor: "#86efac" };
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
      const usuarioBartender: Usuario = { id: 1, nombre: "Bartender" };
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
              {cargando ? "Ingresando..." : "Ingresar al sistema"}
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
  tipo: "total" | "critico" | "agotados" | "optimo" | null;
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
    default:
      return null;
  }

  const totalUnidades = productosFiltrados.reduce((sum, p) => sum + p.stock, 0);

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
            <div className="text-4xl font-bold mb-1" style={{ color }}>{totalUnidades}</div>
            <div className="text-sm text-slate-500 uppercase tracking-wide">Unidades Totales</div>
          </div>
        </div>

        <div className="p-6">
          <table className="w-full">
            <thead className="border-b border-[#2a2a3e]">
              <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wide">
                <th className="pb-4">Producto</th>
                <th className="pb-4">Categoría</th>
                <th className="pb-4 text-right">Stock</th>
                <th className="pb-4 text-right">Mínimo</th>
                <th className="pb-4">Ubicación</th>
              </tr>
            </thead>
            <tbody>
              {productosFiltrados.map((p, idx) => (
                <tr key={p.id} className="border-b border-[#1a1a2a]">
                  <td className="py-3 text-white font-medium">{p.nombre}</td>
                  <td className="py-3"><span className="text-xs px-3 py-1.5 rounded-full" style={{ background: CATEGORIA_BG[p.categoria], color: CATEGORIA_COLOR[p.categoria], border: `1px solid ${CATEGORIA_BORDER[p.categoria]}` }}>{CATEGORIA_EMOJI[p.categoria]} {p.categoria}</span></td>
                  <td className="py-3 text-right text-xl font-bold" style={{ color }}>{p.stock}</td>
                  <td className="py-3 text-right text-slate-500">{p.stock_minimo}</td>
                  <td className="py-3 text-slate-500">{p.ubicacion || "—"}</td>
                </tr>
              ))}
              {productosFiltrados.length === 0 && (
                <tr>
                  <td colSpan={5} className="text-center py-12 text-slate-500">No hay productos en esta categoría</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

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
    valor: productos.filter(p => p.categoria === cat).reduce((sum, p) => sum + (p.uso_promedio || 0), 0),
    stock: productos.filter(p => p.categoria === cat).reduce((sum, p) => sum + p.stock, 0),
  })).filter(item => item.valor > 0 || item.stock > 0);

  const ultimos7Dias = Array.from({ length: 7 }, (_, i) => {
    const fecha = new Date();
    fecha.setDate(fecha.getDate() - i);
    fecha.setHours(0, 0, 0, 0);
    const fechaStr = fecha.toISOString().split('T')[0];
    const movimientosDia = movimientos.filter(m => m.created_at?.split('T')[0] === fechaStr);
    return { fecha: fechaStr, entradas: movimientosDia.filter(m => m.tipo === "entrada").reduce((sum, m) => sum + m.cantidad, 0), salidas: movimientosDia.filter(m => m.tipo === "salida").reduce((sum, m) => sum + m.cantidad, 0) };
  }).reverse();

  const topProductos = productos.map(p => {
    const salidas = movimientos.filter(m => m.producto_id === p.id && m.tipo === "salida").reduce((sum, m) => sum + m.cantidad, 0);
    return { nombre: p.nombre, salidas, stock: p.stock };
  }).sort((a, b) => b.salidas - a.salidas).slice(0, 10);

  const COLORS = ['#fbbf24', '#60a5fa', '#c084fc', '#4ade80', '#f472b6', '#f97316'];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-6" onClick={onClose}>
      <div className="relative w-full max-w-6xl max-h-[90vh] overflow-auto rounded-2xl bg-[#0f1117] border border-[#2a2a3e] p-6 shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="flex justify-between items-center mb-6 sticky top-0 bg-[#0f1117] z-10 pb-4 border-b border-[#2a2a3e]">
          <h2 className="text-2xl font-bold text-white">📊 Estadísticas y Análisis</h2>
          <button onClick={onClose} className="w-10 h-10 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-white text-xl transition-all flex items-center justify-center">✕</button>
        </div>
        
        <div className="mb-8"><h3 className="text-base font-semibold text-slate-300 mb-4">Consumo Promedio por Categoría</h3><div style={{ height: 400 }}><ResponsiveContainer><BarChart data={consumoPorCategoria}><CartesianGrid stroke="#2a2a3e" /><XAxis dataKey="name" stroke="#94a3b8" fontSize={12} /><YAxis stroke="#94a3b8" fontSize={12} /><Tooltip contentStyle={{ background: '#0f1117', border: '1px solid #2a2a3e', borderRadius: '8px' }} /><Legend /><Bar dataKey="valor" name="Consumo Promedio (und/día)" fill="#60a5fa" radius={[8,8,0,0]} /><Bar dataKey="stock" name="Stock Actual" fill="#22c55e" radius={[8,8,0,0]} /></BarChart></ResponsiveContainer></div></div>
        
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
  const [tipoReporte, setTipoReporte] = useState<"todos" | "entradas" | "salidas">("todos");
  const [productoFiltro, setProductoFiltro] = useState<string>("");

  const movimientosFiltrados = movimientos.filter(m => {
    const fechaMov = m.created_at?.split('T')[0];
    const cumpleFecha = fechaMov >= fechaInicio && fechaMov <= fechaFin;
    const cumpleTipo = tipoReporte === "todos" || m.tipo === tipoReporte;
    const cumpleProducto = !productoFiltro || m.producto_id === parseInt(productoFiltro);
    return cumpleFecha && cumpleTipo && cumpleProducto;
  });

  const exportarReporte = async () => {
    const wb = new ExcelJS.Workbook();
    const sh = wb.addWorksheet("Reporte_Movimientos");
    const headerRow = sh.addRow(["Fecha", "Producto", "Tipo", "Cantidad", "Motivo", "Notas"]);
    headerRow.font = { bold: true, size: 12 };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1f2937" } };
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    
    movimientosFiltrados.forEach(m => {
      const prod = productos.find(p => p.id === m.producto_id);
      sh.addRow([new Date(m.created_at).toLocaleDateString("es-CO"), prod?.nombre || "?", m.tipo === "entrada" ? "Ingreso" : "Salida", m.cantidad, m.motivo || "-", m.notas || "-"]);
    });
    
    sh.columns = [{ width: 15 }, { width: 35 }, { width: 12 }, { width: 12 }, { width: 25 }, { width: 35 }];
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `reporte_inventario_${fechaInicio}_a_${fechaFin}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
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
          <div><label className="block text-sm font-medium text-slate-400 mb-2">Tipo de Movimiento</label><select value={tipoReporte} onChange={(e) => setTipoReporte(e.target.value as any)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"><option value="todos">Todos</option><option value="entradas">Solo Entradas</option><option value="salidas">Solo Salidas</option></select></div>
        </div>
        
        <div className="mb-6"><label className="block text-sm font-medium text-slate-400 mb-2">Filtrar por Producto (opcional)</label><select value={productoFiltro} onChange={(e) => setProductoFiltro(e.target.value)} className="w-full px-4 py-2.5 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"><option value="">Todos los productos</option>{productos.map(p => <option key={p.id} value={p.id}>{p.nombre}</option>)}</select></div>
        
        <div className="mb-6 p-4 rounded-xl bg-amber-500/10 border border-amber-500/30"><p className="text-sm text-amber-300">📊 Resumen del reporte: {movimientosFiltrados.length} movimientos encontrados | ↑ Entradas: {movimientosFiltrados.filter(m => m.tipo === "entrada").reduce((s, m) => s + m.cantidad, 0)} unidades | ↓ Salidas: {movimientosFiltrados.filter(m => m.tipo === "salida").reduce((s, m) => s + m.cantidad, 0)} unidades</p></div>
        
        <div className="flex gap-4 justify-end pt-4"><button onClick={onClose} className="px-6 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-white font-medium transition-all">Cancelar</button><button onClick={exportarReporte} className="px-6 py-2.5 rounded-xl bg-gradient-to-r from-amber-600 to-amber-500 hover:from-amber-500 hover:to-amber-600 text-white font-medium transition-all flex items-center gap-2">📥 Exportar a Excel</button></div>
      </div>
    </div>
  );
}

// ==================== VISTA DE INVENTARIO ====================

function VistaInventario({ productos }: { productos: Producto[] }) {
  const [busqueda, setBusqueda] = useState("");
  const [categoriaFiltro, setCategoriaFiltro] = useState<string | null>(null);

  const productosFiltrados = productos.filter(p => 
    p.nombre.toLowerCase().includes(busqueda.toLowerCase()) && 
    (!categoriaFiltro || p.categoria === categoriaFiltro)
  );

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-white mb-2">📦 Inventario General</h2>
        <p className="text-slate-500">Listado completo de todos los productos con sus cantidades actuales</p>
      </div>

      {/* Filtros */}
      <div className="flex gap-3 mb-6 flex-wrap items-center">
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">🔍</span>
          <input type="text" placeholder="Buscar producto..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} className="pl-9 pr-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-72 focus:outline-none focus:border-amber-500 transition-all" />
        </div>
        <button onClick={() => setCategoriaFiltro(null)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all ${!categoriaFiltro ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>Todas</button>
        {CATEGORIAS.map(cat => <button key={cat} onClick={() => setCategoriaFiltro(cat)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all flex items-center gap-2 ${categoriaFiltro === cat ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}><span>{CATEGORIA_EMOJI[cat]}</span>{cat}</button>)}
      </div>

      {/* Tabla de inventario */}
      <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
        <div className="px-5 py-4 border-b border-[#2a2a3e] bg-[#0a0d16]">
          <span className="text-base font-semibold text-white">📋 Lista de Productos</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-[#0a0d16] border-b border-[#2a2a3e]">
              <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wider">
                <th className="p-4">Categoría</th>
                <th className="p-4">Producto</th>
                <th className="p-4">Ubicación</th>
                <th className="p-4">Stock Mínimo</th>
                <th className="p-4 text-center">Stock Actual</th>
                <th className="p-4">Estado</th>
              </tr>
            </thead>
            <tbody>
              {productosFiltrados.map((p, i) => {
                const est = estadoInfo(p);
                return (
                  <tr key={p.id} className={`border-b border-[#1a1a2a] hover:bg-[#1a1a2a] transition-all ${i % 2 === 0 ? 'bg-transparent' : 'bg-[#0a0d16]/50'}`}>
                    <td className="p-4"><span className="text-xs px-3 py-1.5 rounded-full" style={{ background: CATEGORIA_BG[p.categoria], color: CATEGORIA_COLOR[p.categoria], border: `1px solid ${CATEGORIA_BORDER[p.categoria]}` }}>{CATEGORIA_EMOJI[p.categoria]} {p.categoria}</span></td>
                    <td className="p-4"><span className="text-white font-medium text-base">{p.nombre}</span></td>
                    <td className="p-4 text-slate-400 text-sm">{p.ubicacion || "—"}</td>
                    <td className="p-4 text-slate-400 text-sm">{p.stock_minimo} und</td>
                    <td className="p-4 text-center"><span className="text-3xl font-bold" style={{ color: est.color }}>{p.stock}</span><span className="text-sm text-slate-500 ml-1">und</span></td>
                    <td className="p-4"><span className="text-xs px-3 py-1.5 rounded-full font-semibold" style={{ background: est.bg, color: est.color, border: `1px solid ${est.color}40` }}><span className="inline-block w-2 h-2 rounded-full mr-1.5" style={{ background: est.dot }} />{est.label}</span></td>
                  </tr>
                );
              })}
              {productosFiltrados.length === 0 && <tr><td colSpan={6} className="text-center py-12 text-slate-500">No se encontraron productos</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {/* Resumen de stock */}
      <div className="mt-6 grid grid-cols-4 gap-4">
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-blue-400">{productos.length}</div>
          <div className="text-xs text-slate-500">Total Productos</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-green-400">{productos.reduce((sum, p) => sum + p.stock, 0)}</div>
          <div className="text-xs text-slate-500">Unidades Totales</div>
        </div>
        <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-xl p-4 text-center">
          <div className="text-2xl font-bold text-orange-400">{productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length}</div>
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
  const [filtroTipo, setFiltroTipo] = useState<"todos" | "entrada" | "salida">("todos");
  const [filtroProducto, setFiltroProducto] = useState<string>("");
  const [fechaInicio, setFechaInicio] = useState("");
  const [fechaFin, setFechaFin] = useState("");

  const movimientosFiltrados = movimientos.filter(m => {
    const cumpleTipo = filtroTipo === "todos" || m.tipo === filtroTipo;
    const cumpleProducto = !filtroProducto || m.producto_id === parseInt(filtroProducto);
    const fechaMov = new Date(m.created_at).toISOString().split('T')[0];
    const cumpleFechaInicio = !fechaInicio || fechaMov >= fechaInicio;
    const cumpleFechaFin = !fechaFin || fechaMov <= fechaFin;
    return cumpleTipo && cumpleProducto && cumpleFechaInicio && cumpleFechaFin;
  });

  return (
    <div className="p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-white mb-2">🔄 Historial de Movimientos</h2>
        <p className="text-slate-500">Registro completo de todas las entradas y salidas de inventario</p>
      </div>

      {/* Filtros */}
      <div className="grid grid-cols-4 gap-4 mb-6">
        <div>
          <label className="block text-xs font-medium text-slate-400 mb-1">Tipo de Movimiento</label>
          <select value={filtroTipo} onChange={(e) => setFiltroTipo(e.target.value as any)} className="w-full px-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500">
            <option value="todos">Todos</option>
            <option value="entrada">↑ Entradas</option>
            <option value="salida">↓ Salidas</option>
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

      {/* Tabla de movimientos */}
      <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
        <div className="px-5 py-4 border-b border-[#2a2a3e] bg-[#0a0d16]">
          <span className="text-base font-semibold text-white">📋 Registro de Movimientos</span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead className="bg-[#0a0d16] border-b border-[#2a2a3e]">
              <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wider">
                <th className="p-4">Fecha</th>
                <th className="p-4">Producto</th>
                <th className="p-4">Tipo</th>
                <th className="p-4 text-center">Cantidad</th>
                <th className="p-4">Motivo</th>
                <th className="p-4">Notas</th>
              </tr>
            </thead>
            <tbody>
              {movimientosFiltrados.slice(0, 100).map((m, i) => {
                const prod = productos.find(p => p.id === m.producto_id);
                const isEntrada = m.tipo === "entrada";
                return (
                  <tr key={m.id} className={`border-b border-[#1a1a2a] hover:bg-[#1a1a2a] transition-all ${i % 2 === 0 ? 'bg-transparent' : 'bg-[#0a0d16]/50'}`}>
                    <td className="p-4 text-slate-300 text-sm">{new Date(m.created_at).toLocaleDateString("es-CO")}</td>
                    <td className="p-4 text-white font-medium">{prod?.nombre || "?"}</td>
                    <td className="p-4"><span className={`text-xs px-3 py-1 rounded-full font-semibold ${isEntrada ? 'bg-green-500/20 text-green-400 border border-green-500/30' : 'bg-red-500/20 text-red-400 border border-red-500/30'}`}>{isEntrada ? "↑ ENTRADA" : "↓ SALIDA"}</span></td>
                    <td className="p-4 text-center"><span className={`text-xl font-bold ${isEntrada ? 'text-green-400' : 'text-red-400'}`}>{m.cantidad}</span></td>
                    <td className="p-4 text-slate-400 text-sm">{m.motivo || "—"}</td>
                    <td className="p-4 text-slate-500 text-sm">{m.notas || "—"}</td>
                  </tr>
                );
              })}
              {movimientosFiltrados.length === 0 && <tr><td colSpan={6} className="text-center py-12 text-slate-500">No hay movimientos registrados</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {/* Resumen de movimientos */}
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
  const [categoriaFiltro, setCategoriaFiltro] = useState<string | null>(null);
  const [pestanaActiva, setPestanaActiva] = useState<"ingreso" | "salida">("ingreso");
  const [modalAbierto, setModalAbierto] = useState<ModalData["tipo"]>(null);
  const [modalEstadisticas, setModalEstadisticas] = useState(false);
  const [modalReportes, setModalReportes] = useState(false);
  const [filtroMovimientos, setFiltroMovimientos] = useState<"todos" | "entradas" | "salidas">("todos");
  const [vistaActual, setVistaActual] = useState<Vista>("dashboard");

  const [productoEntrada, setProductoEntrada] = useState("");
  const [cantidadEntrada, setCantidadEntrada] = useState("");
  const [fechaEntrada, setFechaEntrada] = useState(new Date().toISOString().split('T')[0]);
  const [notasEntrada, setNotasEntrada] = useState("");
  const [loadingEntrada, setLoadingEntrada] = useState(false);
  const [productoSalida, setProductoSalida] = useState("");
  const [cantidadSalida, setCantidadSalida] = useState("");
  const [notasSalida, setNotasSalida] = useState("");
  const [loadingSalida, setLoadingSalida] = useState(false);

  useEffect(() => {
    const guardado = localStorage.getItem("usuario_activo");
    if (guardado) { try { const u = JSON.parse(guardado); if (u && u.nombre && u.id) setUsuario(u); } catch {} }
  }, []);

  const cargarProductos = async () => { const { data } = await supabase.from("productos").select("*").order("nombre"); if (data) setProductos(data); };
  const cargarMovimientos = async () => { const { data } = await supabase.from("movimientos").select("*").order("created_at", { ascending: false }).limit(200); if (data) setMovimientos(data); };

  useEffect(() => { if (usuario) { (async () => { await cargarProductos(); await cargarMovimientos(); setCargando(false); })(); } }, [usuario]);

  const productosFiltrados = productos.filter(p => p.nombre.toLowerCase().includes(busqueda.toLowerCase()) && (!categoriaFiltro || p.categoria === categoriaFiltro));
  const totalProductos = productos.length;
  const stockCritico = productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length;
  const sinStock = productos.filter(p => p.stock === 0).length;
  const stockOptimo = productos.filter(p => p.stock > p.stock_minimo * 1.5).length;
  const entradasHoy = movimientos.filter(m => m.tipo === "entrada" && new Date(m.created_at).toDateString() === new Date().toDateString()).reduce((s, m) => s + m.cantidad, 0);
  const salidasHoy = movimientos.filter(m => m.tipo === "salida" && new Date(m.created_at).toDateString() === new Date().toDateString()).reduce((s, m) => s + m.cantidad, 0);
  const movimientosFiltradosHome = movimientos.filter(m => filtroMovimientos === "entradas" ? m.tipo === "entrada" : filtroMovimientos === "salidas" ? m.tipo === "salida" : true);

  const registrarEntrada = async () => {
    if (!productoEntrada || !cantidadEntrada) { alert("Completa todos los campos"); return; }
    const prod = productos.find(p => p.nombre === productoEntrada);
    if (!prod) return;
    setLoadingEntrada(true);
    const cantidad = parseInt(cantidadEntrada);
    await supabase.from("productos").update({ stock: prod.stock + cantidad }).eq("id", prod.id);
    await supabase.from("movimientos").insert({ producto_id: prod.id, tipo: "entrada", cantidad, motivo: "Compra", fecha: fechaEntrada, notas: notasEntrada });
    setProductoEntrada(""); setCantidadEntrada(""); setNotasEntrada(""); setFechaEntrada(new Date().toISOString().split('T')[0]);
    await cargarProductos(); await cargarMovimientos();
    setLoadingEntrada(false);
    alert("Ingreso registrado correctamente");
  };

  const registrarSalida = async () => {
    if (!productoSalida || !cantidadSalida) { alert("Completa todos los campos"); return; }
    const prod = productos.find(p => p.nombre === productoSalida);
    if (!prod) return;
    const cantidad = parseInt(cantidadSalida);
    if (cantidad > prod.stock) { alert("Stock insuficiente"); return; }
    setLoadingSalida(true);
    await supabase.from("productos").update({ stock: prod.stock - cantidad }).eq("id", prod.id);
    await supabase.from("movimientos").insert({ producto_id: prod.id, tipo: "salida", cantidad, motivo: "Venta", notas: notasSalida });
    setProductoSalida(""); setCantidadSalida(""); setNotasSalida("");
    await cargarProductos(); await cargarMovimientos();
    setLoadingSalida(false);
    alert("Salida registrada correctamente");
  };

  const descargarExcel = async () => {
    const wb = new ExcelJS.Workbook();
    const sh = wb.addWorksheet("Inventario_Completo");
    const headerRow = sh.addRow(["Producto", "Categoría", "Ubicación", "Stock Actual", "Stock Mínimo", "Uso Promedio", "Estado"]);
    headerRow.font = { bold: true, size: 12 };
    headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1f2937" } };
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    
    productos.forEach((p) => {
      const { label } = estadoInfo(p);
      sh.addRow([p.nombre, p.categoria, p.ubicacion || "", p.stock, p.stock_minimo, p.uso_promedio ? `${p.uso_promedio} und/día` : "", label]);
    });
    
    sh.columns = [{ width: 30 }, { width: 18 }, { width: 15 }, { width: 12 }, { width: 12 }, { width: 18 }, { width: 12 }];
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `inventario_${new Date().toLocaleDateString("es-CO").replace(/\//g, "-")}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const now = new Date();
  const hora = now.toLocaleTimeString("es-CO", { hour: "2-digit", minute: "2-digit" });
  const fecha = now.toLocaleDateString("es-CO", { day: "numeric", month: "long", year: "numeric" });
  const handleCardClick = (tipo: ModalData["tipo"]) => setModalAbierto(tipo);
  const handleLogout = () => { localStorage.removeItem("usuario_activo"); setUsuario(null); };

  if (!usuario) return <LoginScreen onLogin={setUsuario} />;
  if (cargando) return <div className="min-h-screen bg-[#0a0d16] flex items-center justify-center"><div className="text-center"><div className="animate-spin rounded-full h-12 w-12 border-4 border-amber-500 border-t-transparent mx-auto mb-4"></div><p className="text-slate-400">Cargando sistema...</p></div></div>;

  // Renderizar contenido según la vista actual
  const renderContenido = () => {
    switch (vistaActual) {
      case "inventario":
        return <VistaInventario productos={productos} />;
      case "movimientos":
        return <VistaMovimientos movimientos={movimientos} productos={productos} />;
      default:
        return (
          <div className="p-6">
            {/* KPI Cards */}
            <div className="grid grid-cols-4 gap-5 mb-8">
              {[
                { label: "Total Productos", value: totalProductos, sub: "Productos registrados", color: "#3b82f6", bg: "#1e3a5f", icon: "📦", border: "#3b82f6", tipo: "total" as const },
                { label: "Stock Crítico", value: stockCritico, sub: "Requieren reposición", color: "#f97316", bg: "#3a1e0a", icon: "⚠️", border: "#f97316", tipo: "critico" as const },
                { label: "Agotados", value: sinStock, sub: "Sin stock disponible", color: "#ef4444", bg: "#3a0a0a", icon: "🚫", border: "#ef4444", tipo: "agotados" as const },
                { label: "Stock Óptimo", value: stockOptimo, sub: "Nivel adecuado", color: "#22c55e", bg: "#0a3a1e", icon: "✅", border: "#22c55e", tipo: "optimo" as const },
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

            {/* Filters */}
            <div className="flex gap-3 mb-6 flex-wrap items-center">
              <div className="relative">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500">🔍</span>
                <input type="text" placeholder="Buscar producto..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} className="pl-9 pr-4 py-2.5 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-72 focus:outline-none focus:border-amber-500 transition-all" />
              </div>
              <button onClick={() => setCategoriaFiltro(null)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all ${!categoriaFiltro ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}>Todas las categorías</button>
              {CATEGORIAS.map(cat => <button key={cat} onClick={() => setCategoriaFiltro(cat)} className={`px-4 py-2 rounded-xl text-sm font-medium transition-all flex items-center gap-2 ${categoriaFiltro === cat ? 'bg-gradient-to-r from-amber-600 to-amber-500 text-white shadow-md' : 'bg-[#0f1117] text-slate-400 border border-[#2a2a3e] hover:bg-slate-800'}`}><span>{CATEGORIA_EMOJI[cat]}</span>{cat}</button>)}
            </div>

            {/* Products Table */}
            <div className="bg-[#0f1117] border border-[#2a2a3e] rounded-2xl overflow-hidden shadow-xl">
              <div className="px-5 py-4 flex items-center justify-between border-b border-[#2a2a3e] bg-[#0a0d16]">
                <span className="text-base font-semibold text-white">📋 Productos del inventario</span>
                <div className="relative">
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-sm">🔍</span>
                  <input placeholder="Filtrar productos..." value={busqueda} onChange={(e) => setBusqueda(e.target.value)} className="pl-9 pr-4 py-2 rounded-xl bg-[#0f1117] border border-[#2a2a3e] text-white text-sm w-64 focus:outline-none focus:border-amber-500" />
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[1000px]">
                  <thead className="bg-[#0a0d16] border-b border-[#2a2a3e]">
                    <tr className="text-left text-xs font-semibold text-slate-500 uppercase tracking-wider">
                      <th className="p-4">Categoría</th>
                      <th className="p-4">Producto</th>
                      <th className="p-4">Ubicación</th>
                      <th className="p-4">Uso Promedio</th>
                      <th className="p-4">Stock Actual</th>
                      <th className="p-4">Estado</th>
                      <th className="p-4">Acción</th>
                    </tr>
                  </thead>
                  <tbody>
                    {productosFiltrados.map((p, i) => {
                      const est = estadoInfo(p);
                      return (
                        <tr key={p.id} className={`border-b border-[#1a1a2a] hover:bg-[#1a1a2a] transition-all ${i % 2 === 0 ? 'bg-transparent' : 'bg-[#0a0d16]/50'}`}>
                          <td className="p-4"><span className="text-xs px-3 py-1.5 rounded-full" style={{ background: CATEGORIA_BG[p.categoria], color: CATEGORIA_COLOR[p.categoria], border: `1px solid ${CATEGORIA_BORDER[p.categoria]}` }}>{CATEGORIA_EMOJI[p.categoria]} {p.categoria}</span></td>
                          <td className="p-4"><span className="text-white font-medium text-base">{p.nombre}</span></td>
                          <td className="p-4 text-slate-400 text-sm">{p.ubicacion || "—"}</td>
                          <td className="p-4 text-slate-400 text-sm">{p.uso_promedio ? `${p.uso_promedio} und/día` : "—"}</td>
                          <td className="p-4"><span className="text-2xl font-bold" style={{ color: est.color }}>{p.stock}</span><span className="text-sm text-slate-500 ml-1">und</span></td>
                          <td className="p-4"><span className="text-xs px-3 py-1.5 rounded-full font-semibold" style={{ background: est.bg, color: est.color, border: `1px solid ${est.color}40` }}><span className="inline-block w-2 h-2 rounded-full mr-1.5" style={{ background: est.dot }} />{est.label}</span></td>
                          <td className="p-4"><button onClick={() => { setPestanaActiva("salida"); setProductoSalida(p.nombre); }} className="px-4 py-2 rounded-xl bg-red-500/10 border border-red-500/30 text-red-400 text-sm font-medium hover:bg-red-500/20 transition-all">Registrar Salida</button></td>
                        </tr>
                      );
                    })}
                    {productosFiltrados.length === 0 && <tr><td colSpan={7} className="text-center py-12 text-slate-500">No se encontraron productos</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Recent Movements */}
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
      {/* HEADER */}
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
        {/* SIDEBAR */}
        <aside className="w-64 bg-[#0f1117] border-r border-[#2a2a3e] flex flex-col flex-shrink-0">
          <div className="p-5 border-b border-[#2a2a3e]">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-xl bg-amber-500/20 flex items-center justify-center text-lg">🍺</div>
              <div>
                <div className="text-sm font-semibold text-white">INVENTARIO</div>
                <div className="text-[10px] text-slate-500">Version 1.0</div>
              </div>
            </div>
          </div>
          <nav className="flex-1 p-3 space-y-1">
            <div className="px-3 py-2 text-[10px] text-slate-500 uppercase tracking-wider font-semibold">Principal</div>
            <div onClick={() => setVistaActual("dashboard")} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer ${vistaActual === "dashboard" ? 'bg-gradient-to-r from-amber-500/20 to-amber-500/10 text-amber-400' : 'text-slate-400 hover:bg-slate-800'}`}><span>📊</span> Dashboard</div>
            <div onClick={() => setVistaActual("inventario")} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer ${vistaActual === "inventario" ? 'bg-gradient-to-r from-amber-500/20 to-amber-500/10 text-amber-400' : 'text-slate-400 hover:bg-slate-800'}`}><span>📦</span> Inventario <span className="ml-auto bg-red-500 text-white text-[10px] rounded-full px-2 py-0.5">{productos.filter(p => p.stock <= p.stock_minimo).length}</span></div>
            <div onClick={() => setVistaActual("movimientos")} className={`flex items-center gap-3 px-4 py-2.5 rounded-xl text-sm font-medium transition-all cursor-pointer ${vistaActual === "movimientos" ? 'bg-gradient-to-r from-amber-500/20 to-amber-500/10 text-amber-400' : 'text-slate-400 hover:bg-slate-800'}`}><span>🔄</span> Movimientos</div>
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

        {/* MAIN CONTENT - Cambia según la vista */}
        <div className="flex-1 overflow-auto">
          {renderContenido()}
        </div>

        {/* RIGHT PANEL - Solo se muestra en dashboard */}
        {vistaActual === "dashboard" && (
          <aside className="w-96 bg-[#0f1117] border-l border-[#2a2a3e] flex flex-col flex-shrink-0 shadow-xl">
            {/* Alertas */}
            <div className="m-5 p-4 rounded-xl bg-gradient-to-r from-orange-500/10 to-orange-600/5 border border-orange-500/30">
              <div className="flex items-center gap-2 text-sm font-semibold text-orange-400 mb-3"><span>⚠️</span> Alertas de stock crítico</div>
              <div className="space-y-2">
                {productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).slice(0, 5).map(p => (
                  <div key={p.id} className="flex justify-between items-center py-2 px-2 rounded-lg bg-slate-800/50">
                    <span className="text-sm text-slate-300">{p.nombre}</span>
                    <span className="text-orange-400 font-bold text-base">{p.stock} und</span>
                  </div>
                ))}
                {productos.filter(p => p.stock <= p.stock_minimo && p.stock > 0).length === 0 && <div className="text-center py-4 text-slate-500 text-sm">No hay alertas activas</div>}
              </div>
            </div>

            {/* Tabs */}
            <div className="flex mx-5 bg-[#0a0d16] rounded-xl p-1 gap-1">
              <button onClick={() => setPestanaActiva("ingreso")} className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${pestanaActiva === "ingreso" ? 'bg-gradient-to-r from-green-600 to-green-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}><span>↑</span> Registrar Ingreso</button>
              <button onClick={() => setPestanaActiva("salida")} className={`flex-1 py-3 rounded-lg text-sm font-semibold transition-all flex items-center justify-center gap-2 ${pestanaActiva === "salida" ? 'bg-gradient-to-r from-red-600 to-red-500 text-white shadow-md' : 'text-slate-400 hover:bg-slate-800'}`}><span>↓</span> Registrar Salida</button>
            </div>

            {/* Forms */}
            <div className="flex-1 overflow-auto p-5">
              {pestanaActiva === "ingreso" ? (
                <div className="space-y-4">
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Producto</label><select value={productoEntrada} onChange={(e) => setProductoEntrada(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500"><option value="">Seleccionar producto...</option>{productos.map(p => <option key={p.id} value={p.nombre}>{p.nombre} — Stock actual: {p.stock}</option>)}</select></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Cantidad</label><input type="number" value={cantidadEntrada} onChange={(e) => setCantidadEntrada(e.target.value)} placeholder="Número de unidades" className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" /></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Fecha de ingreso</label><input type="date" value={fechaEntrada} onChange={(e) => setFechaEntrada(e.target.value)} className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm focus:outline-none focus:border-amber-500" /></div>
                  <div><label className="block text-sm font-medium text-slate-400 mb-2">Notas</label><textarea value={notasEntrada} onChange={(e) => setNotasEntrada(e.target.value)} rows={3} placeholder="Observaciones adicionales..." className="w-full px-4 py-3 rounded-xl bg-[#0a0d16] border border-[#2a2a3e] text-white text-sm resize-none focus:outline-none focus:border-amber-500" /></div>
                  <button onClick={registrarEntrada} disabled={loadingEntrada || !productoEntrada || !cantidadEntrada} className="w-full py-3 rounded-xl bg-gradient-to-r from-green-600 to-green-500 hover:from-green-500 hover:to-green-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2"><span>↑</span> {loadingEntrada ? "PROCESANDO..." : "REGISTRAR INGRESO"}</button>
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
                  <button onClick={registrarSalida} disabled={loadingSalida || !productoSalida || !cantidadSalida} className="w-full py-3 rounded-xl bg-gradient-to-r from-red-600 to-red-500 hover:from-red-500 hover:to-red-600 text-white font-semibold transition-all disabled:opacity-50 flex items-center justify-center gap-2"><span>↓</span> {loadingSalida ? "PROCESANDO..." : "REGISTRAR SALIDA"}</button>
                </div>
              )}
            </div>
          </aside>
        )}
      </div>

      {/* MODALES */}
      {modalAbierto && <ModalDetalle tipo={modalAbierto} productos={productos} onClose={() => setModalAbierto(null)} />}
      {modalEstadisticas && <ModalEstadisticas productos={productos} movimientos={movimientos} onClose={() => setModalEstadisticas(false)} />}
      {modalReportes && <ModalReportes productos={productos} movimientos={movimientos} onClose={() => setModalReportes(false)} />}
    </div>
  );
}