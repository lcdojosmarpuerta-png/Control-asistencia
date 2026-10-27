/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Search,
  Key,
  UserCheck,
  LogIn,
  Coffee,
  Utensils,
  LogOut,
  History,
  RotateCcw,
  FileSpreadsheet,
  Printer,
  Inbox,
  UploadCloud,
  UserPlus,
  Download,
  FileUp,
  Users,
  ClockAlert,
  RefreshCw,
  BarChart2,
  PieChart,
  ListFilter,
  CheckCircle,
  X,
  Lock,
  ShieldCheck,
  KeyRound,
  Fingerprint,
  Database,
  BarChart3,
  Wrench,
  Trash2,
  Edit3,
  Sparkles,
  Clock,
  ShieldAlert,
  Timer
} from 'lucide-react';

export interface Colaborador {
  id: string;
  locker: string;
  nombre: string;
  area: string;
  turnoBase?: 'auto' | '07:00' | '08:00' | '09:00';
}

export type TipoMarcaje =
  | 'Entrada'
  | 'Salida a Descanso'
  | 'Retorno de Descanso'
  | 'Salida';

export interface Marcaje {
  id: number;
  colabId: string;
  timestamp: string;
  fecha: string;
  hora: string;
  locker: string;
  nombre: string;
  area: string;
  tipo: TipoMarcaje;
}

interface ToastItem {
  id: number;
  msg: string;
  type: 'success' | 'error' | 'warning' | 'info';
}

const CLAVE_STORAGE_KEY = 'asistencia_clave_supervisor';
const PERSONAL_STORAGE_KEY = 'asistencia_personal_mc';
const MARCAJES_STORAGE_KEY = 'asistencia_marcajes_mc';
const PERFIL_STORAGE_KEY = 'asistencia_perfil_rol';
const MIN_SEGUNDOS_ENTRE_MARCAJES = 60; // 1 minuto mínimo para evitar registros accidentales

const TURNOS_ENTRADA = [
  { codigo: '07:00', nombre: '07:00 AM', minutosBase: 7 * 60, toleranciaMax: 7 * 60 + 5 },
  { codigo: '08:00', nombre: '08:00 AM', minutosBase: 8 * 60, toleranciaMax: 8 * 60 + 5 },
  { codigo: '09:00', nombre: '09:00 AM', minutosBase: 9 * 60, toleranciaMax: 9 * 60 + 5 }
];

const INITIAL_PERSONAL: Colaborador[] = [
  { id: '101', locker: '101', nombre: 'Carlos Eduardo Mendoza', area: 'Piso de Venta / Ferretería', turnoBase: '07:00' },
  { id: '102', locker: '102', nombre: 'Mariana Sofia Silva', area: 'Cajas y Facturación', turnoBase: '08:00' },
  { id: '103', locker: '103', nombre: 'José Gregorio Ramos', area: 'Almacén Central / Despacho', turnoBase: '07:00' },
  { id: '104', locker: '104', nombre: 'Elena Beatriz Rivas', area: 'Atención al Cliente', turnoBase: '09:00' },
  { id: '105', locker: '105', nombre: 'Luis Alberto Páez', area: 'Seguridad / Control', turnoBase: '07:00' },
  { id: '106', locker: 'S/N', nombre: 'Gabriela Fernanda Torres', area: 'Administración', turnoBase: '08:00' }
];

function cleanStr(str: string | undefined | null): string {
  return String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function parseHoraToSeconds(horaStr: string): number | null {
  if (!horaStr || typeof horaStr !== 'string') return null;
  const parts = horaStr.split(':').map((n) => parseInt(n, 10));
  if (parts.length < 2 || isNaN(parts[0]) || isNaN(parts[1])) return null;
  const h = parts[0];
  const m = parts[1];
  const s = parts[2] && !isNaN(parts[2]) ? parts[2] : 0;
  return h * 3600 + m * 60 + s;
}

function formatearTiempoTranscurrido(segundosTotales: number): string {
  if (segundosTotales < 0) return '0 seg';
  if (segundosTotales < 60) {
    return `${segundosTotales} seg`;
  }
  const horas = Math.floor(segundosTotales / 3600);
  const minutos = Math.floor((segundosTotales % 3600) / 60);
  const segundos = segundosTotales % 60;

  if (horas > 0) {
    return `${horas}h ${minutos}m ${segundos}s`;
  }
  return `${minutos} min ${segundos.toString().padStart(2, '0')} seg`;
}

function calcularRetardoEntrada(
  horaStr: string,
  turnoSeleccionado: string = 'todos',
  turnoBaseColaborador?: string
) {
  if (!horaStr || typeof horaStr !== 'string' || horaStr === '--:--') return null;
  const partes = horaStr.split(':');
  if (partes.length < 2) return null;

  const horas = parseInt(partes[0], 10);
  const minutos = parseInt(partes[1], 10);
  const segundos = partes[2] ? parseInt(partes[2], 10) : 0;
  if (isNaN(horas) || isNaN(minutos)) return null;

  const minutosReales = horas * 60 + minutos + segundos / 60;

  let turnoAsignado = TURNOS_ENTRADA[0];

  if (turnoSeleccionado === '07:00') {
    turnoAsignado = TURNOS_ENTRADA[0];
  } else if (turnoSeleccionado === '08:00') {
    turnoAsignado = TURNOS_ENTRADA[1];
  } else if (turnoSeleccionado === '09:00') {
    turnoAsignado = TURNOS_ENTRADA[2];
  } else if (turnoBaseColaborador && turnoBaseColaborador !== 'auto') {
    const found = TURNOS_ENTRADA.find((t) => t.codigo === turnoBaseColaborador);
    if (found) turnoAsignado = found;
  } else {
    if (minutosReales <= 7 * 60 + 30) {
      turnoAsignado = TURNOS_ENTRADA[0];
    } else if (minutosReales <= 8 * 60 + 30) {
      turnoAsignado = TURNOS_ENTRADA[1];
    } else {
      turnoAsignado = TURNOS_ENTRADA[2];
    }
  }

  const diffMinutos = Math.round(minutosReales - turnoAsignado.minutosBase);
  const tieneRetardoCritico = minutosReales > turnoAsignado.toleranciaMax;

  return {
    turno: turnoAsignado.nombre,
    codigoTurno: turnoAsignado.codigo,
    horaReal: horaStr,
    minutosRetardo: diffMinutos > 0 ? diffMinutos : 0,
    esRetardoCritico: tieneRetardoCritico
  };
}

export default function App() {
  // Reloj en vivo
  const [now, setNow] = useState<Date>(() => new Date());

  // Estado principal con persistencia LocalStorage compatible
  const [personal, setPersonal] = useState<Colaborador[]>(() => {
    try {
      const saved =
        localStorage.getItem(PERSONAL_STORAGE_KEY) ||
        localStorage.getItem('asistencia_personal');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed.map((colab: Colaborador, idx: number) => ({
            ...colab,
            id:
              colab.id ||
              (colab.locker && colab.locker !== 'S/N'
                ? String(colab.locker).trim()
                : `COLAB_${Date.now()}_${idx}`)
          }));
        }
      }
    } catch {
      // Fallback
    }
    return INITIAL_PERSONAL;
  });

  const [marcajes, setMarcajes] = useState<Marcaje[]>(() => {
    try {
      const saved =
        localStorage.getItem(MARCAJES_STORAGE_KEY) ||
        localStorage.getItem('asistencia_marcajes');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {
      // Fallback
    }
    return [];
  });

  const [perfilActual, setPerfilActual] = useState<'auxiliar' | 'supervisor'>(() => {
    const saved = localStorage.getItem(PERFIL_STORAGE_KEY);
    return saved === 'supervisor' ? 'supervisor' : 'auxiliar';
  });

  const [activeTab, setActiveTab] = useState<'terminal' | 'reportes' | 'config' | 'graficos'>('terminal');

  // Terminal states
  const [lockerQuery, setLockerQuery] = useState('');
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [colaboradorActual, setColaboradorActual] = useState<Colaborador | null>(null);
  const [liveSearchQuery, setLiveSearchQuery] = useState('');
  const [customClockTime, setCustomClockTime] = useState<string>(''); // Solo para supervisor si desea registrar hora manual o probar

  // Alerta visual de intento de marcaje accidental (< 1 minuto)
  const [alertaMarcajeAccidental, setAlertaMarcajeAccidental] = useState<{
    intentoTipo: TipoMarcaje;
    ultimoTipo: TipoMarcaje;
    ultimaHora: string;
    timestampIntento: number;
  } | null>(null);

  // Reportes states
  const [filtroPeriodo, setFiltroPeriodo] = useState<'diario' | 'semanal' | 'quincenal' | 'mensual' | 'todos'>('diario');
  const [filtroArea, setFiltroArea] = useState('');
  const [filtroBusqueda, setFiltroBusqueda] = useState('');

  // Gráficos states
  const [filtroGraficoPeriodo, setFiltroGraficoPeriodo] = useState<'diario' | 'semanal' | 'quincenal' | 'mensual' | 'todos'>('semanal');
  const [filtroGraficoTurno, setFiltroGraficoTurno] = useState<'todos' | '07:00' | '08:00' | '09:00'>('todos');
  const [filtroGraficoArea, setFiltroGraficoArea] = useState('');

  // Modales states
  const [modalNuevoPersonalOpen, setModalNuevoPersonalOpen] = useState(false);
  const [editingColab, setEditingColab] = useState<Colaborador | null>(null);
  const [nuevoNombre, setNuevoNombre] = useState('');
  const [nuevaArea, setNuevaArea] = useState('');
  const [nuevoLocker, setNuevoLocker] = useState('');
  const [nuevoTurnoBase, setNuevoTurnoBase] = useState<'auto' | '07:00' | '08:00' | '09:00'>('auto');

  const [modalLoginSupervisorOpen, setModalLoginSupervisorOpen] = useState(false);
  const [claveSupervisorInput, setClaveSupervisorInput] = useState('');

  const [modalCambiarClaveOpen, setModalCambiarClaveOpen] = useState(false);
  const [claveActualInput, setClaveActualInput] = useState('');
  const [nuevaClaveInput, setNuevaClaveInput] = useState('');

  const [confirmDialog, setConfirmDialog] = useState<{
    open: boolean;
    title: string;
    message: string;
    confirmLabel: string;
    onConfirm: () => void;
  } | null>(null);

  // Toasts
  const [toasts, setToasts] = useState<ToastItem[]>([]);

  const searchInputRef = useRef<HTMLInputElement>(null);
  const suggestionsBoxRef = useRef<HTMLDivElement>(null);

  const showToast = (msg: string, type: ToastItem['type'] = 'info') => {
    const id = Date.now() + Math.random();
    setToasts((prev) => [...prev, { id, msg, type }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 4200);
  };

  // Reloj 1s
  useEffect(() => {
    const interval = setInterval(() => {
      setNow(new Date());
    }, 1000);
    return () => clearInterval(interval);
  }, []);

  // Sincronización LocalStorage
  useEffect(() => {
    localStorage.setItem(PERSONAL_STORAGE_KEY, JSON.stringify(personal));
  }, [personal]);

  useEffect(() => {
    localStorage.setItem(MARCAJES_STORAGE_KEY, JSON.stringify(marcajes));
  }, [marcajes]);

  useEffect(() => {
    localStorage.setItem(PERFIL_STORAGE_KEY, perfilActual);
  }, [perfilActual]);

  // Cerrar sugerencias al hacer clic fuera
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (
        suggestionsBoxRef.current &&
        !suggestionsBoxRef.current.contains(e.target as Node) &&
        searchInputRef.current &&
        !searchInputRef.current.contains(e.target as Node)
      ) {
        setShowSuggestions(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const obtenerClaveSupervisor = () => {
    return localStorage.getItem(CLAVE_STORAGE_KEY) || '1234';
  };

  const hoyStr = useMemo(() => now.toLocaleDateString('es-ES'), [now]);
  const timeStr = useMemo(() => now.toLocaleTimeString('es-ES', { hour12: false }), [now]);
  const dateStr = useMemo(
    () =>
      now.toLocaleDateString('es-ES', {
        weekday: 'long',
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      }),
    [now]
  );

  // Áreas únicas
  const areasDisponibles = useMemo(() => {
    return Array.from(new Set(personal.map((p) => p.area).filter(Boolean))).sort();
  }, [personal]);

  // Sugerencias de búsqueda en vivo
  const suggestions = useMemo(() => {
    const q = cleanStr(lockerQuery);
    if (!q) return [];
    return personal
      .filter((p) => {
        const matchLocker = p.locker && p.locker !== 'S/N' && cleanStr(p.locker).includes(q);
        const matchNombre = cleanStr(p.nombre).includes(q);
        const matchArea = cleanStr(p.area).includes(q);
        return matchLocker || matchNombre || matchArea;
      })
      .slice(0, 6);
  }, [lockerQuery, personal]);

  // Marcajes de un empleado en una fecha
  const getMarcajesEmpleadoFecha = (colab: Colaborador | null, fechaStr: string) => {
    if (!colab) return [];
    return marcajes.filter((m) => {
      if (m.fecha !== fechaStr) return false;
      if (colab.id && m.colabId && colab.id === m.colabId) return true;
      if (colab.locker && colab.locker !== 'S/N' && m.locker && m.locker !== 'S/N') {
        return String(m.locker).trim().toUpperCase() === String(colab.locker).trim().toUpperCase();
      }
      return m.nombre.trim().toLowerCase() === colab.nombre.trim().toLowerCase();
    });
  };

  const registrosHoyColaborador = useMemo(() => {
    return getMarcajesEmpleadoFecha(colaboradorActual, hoyStr);
  }, [colaboradorActual, marcajes, hoyStr]);

  // Último marcaje y cálculo en tiempo real del tiempo transcurrido
  const infoUltimoMarcaje = useMemo(() => {
    if (!colaboradorActual || registrosHoyColaborador.length === 0) return null;

    // Ordenar por timestamp descendente para obtener el último marcaje real efectuado hoy
    const ordenados = [...registrosHoyColaborador].sort(
      (a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime()
    );
    const ultimo = ordenados[0];

    let segundosTranscurridos = Math.max(
      0,
      Math.floor((now.getTime() - new Date(ultimo.timestamp).getTime()) / 1000)
    );

    // Si el supervisor tiene puesta una Hora Test manual, calcular diferencia también respecto a esa hora si aplica
    if (perfilActual === 'supervisor' && customClockTime) {
      const segsTest = parseHoraToSeconds(`${customClockTime}:00`);
      const segsUltimo = parseHoraToSeconds(ultimo.hora);
      if (segsTest !== null && segsUltimo !== null && segsTest >= segsUltimo) {
        segundosTranscurridos = segsTest - segsUltimo;
      }
    }

    const menosDeUnMinuto = segundosTranscurridos < MIN_SEGUNDOS_ENTRE_MARCAJES;
    const segundosRestantes = Math.max(0, MIN_SEGUNDOS_ENTRE_MARCAJES - segundosTranscurridos);
    const porcentajeProgreso = Math.min(
      100,
      Math.round((segundosTranscurridos / MIN_SEGUNDOS_ENTRE_MARCAJES) * 100)
    );

    return {
      ultimo,
      segundosTranscurridos,
      tiempoFormateado: formatearTiempoTranscurrido(segundosTranscurridos),
      menosDeUnMinuto,
      segundosRestantes,
      porcentajeProgreso
    };
  }, [colaboradorActual, registrosHoyColaborador, now, perfilActual, customClockTime]);

  // Limpiar automáticamente la alerta de bloqueo accidental cuando ya se cumple el minuto o cambia de colaborador
  useEffect(() => {
    if (alertaMarcajeAccidental && (!infoUltimoMarcaje || !infoUltimoMarcaje.menosDeUnMinuto)) {
      setAlertaMarcajeAccidental(null);
    }
  }, [infoUltimoMarcaje, alertaMarcajeAccidental]);

  const estadoMarcajesColaborador = useMemo(() => {
    const entrada = registrosHoyColaborador.find((m) => m.tipo === 'Entrada');
    const salidaDescanso = registrosHoyColaborador.find((m) => m.tipo === 'Salida a Descanso');
    const retornoDescanso = registrosHoyColaborador.find((m) => m.tipo === 'Retorno de Descanso');
    const salida = registrosHoyColaborador.find((m) => m.tipo === 'Salida');

    const total = [entrada, salidaDescanso, retornoDescanso, salida].filter(Boolean).length;

    return {
      entrada,
      salidaDescanso,
      retornoDescanso,
      salida,
      total,
      puedeEntrada: !entrada,
      puedeSalDescanso: Boolean(entrada) && !salidaDescanso,
      puedeRetDescanso: Boolean(salidaDescanso) && !retornoDescanso,
      puedeSalida: Boolean(retornoDescanso) && !salida
    };
  }, [registrosHoyColaborador]);

  // Acciones de búsqueda y selección
  const seleccionarColaborador = (colab: Colaborador) => {
    setColaboradorActual(colab);
    setLockerQuery(colab.nombre);
    setShowSuggestions(false);
    setAlertaMarcajeAccidental(null);
  };

  const handleLockerSearch = (e: React.FormEvent) => {
    e.preventDefault();
    const rawVal = lockerQuery.trim();
    const code = cleanStr(rawVal);
    if (!code) return;

    setShowSuggestions(false);
    setAlertaMarcajeAccidental(null);

    let found = personal.find(
      (item) => item.locker && item.locker !== 'S/N' && cleanStr(item.locker) === code
    );
    if (!found) {
      found = personal.find((item) => cleanStr(item.nombre) === code);
    }
    if (!found) {
      const matches = personal.filter(
        (item) =>
          cleanStr(item.nombre).includes(code) ||
          (item.locker && item.locker !== 'S/N' && cleanStr(item.locker).includes(code))
      );
      if (matches.length === 1) {
        found = matches[0];
      } else if (matches.length > 1) {
        setShowSuggestions(true);
        showToast(
          `Se encontraron ${matches.length} coincidencias. Seleccione una de la lista.`,
          'info'
        );
        return;
      }
    }

    if (found) {
      setColaboradorActual(found);
    } else {
      setColaboradorActual(null);
      showToast(`No se encontró personal con: "${rawVal}"`, 'error');
    }
  };

  // Ejecución efectiva del marcaje
  const ejecutarGuardadoMarcaje = (tipo: TipoMarcaje) => {
    if (!colaboradorActual) return;

    const currentMoment = new Date();
    const horaRegistro =
      perfilActual === 'supervisor' && customClockTime
        ? `${customClockTime}:00`
        : currentMoment.toLocaleTimeString('es-ES', { hour12: false });

    const nuevoRegistro: Marcaje = {
      id: Date.now(),
      colabId: colaboradorActual.id,
      timestamp: currentMoment.toISOString(),
      fecha: hoyStr,
      hora: horaRegistro,
      locker:
        colaboradorActual.locker && colaboradorActual.locker !== 'S/N'
          ? colaboradorActual.locker
          : 'S/N',
      nombre: colaboradorActual.nombre,
      area: colaboradorActual.area,
      tipo
    };

    setMarcajes((prev) => [nuevoRegistro, ...prev]);
    setAlertaMarcajeAccidental(null);
    showToast(`¡${tipo} registrada para ${colaboradorActual.nombre} (${horaRegistro})!`, 'success');
  };

  // Registrar Marcaje con Secuencia Estricta (sin bloqueo de espera)
  const registrarMarcaje = (tipo: TipoMarcaje) => {
    if (!colaboradorActual) {
      showToast('Debe consultar un colaborador antes de marcar', 'warning');
      return;
    }

    const { entrada, salidaDescanso, retornoDescanso, salida } = estadoMarcajesColaborador;

    const mapActual: Record<TipoMarcaje, Marcaje | undefined> = {
      Entrada: entrada,
      'Salida a Descanso': salidaDescanso,
      'Retorno de Descanso': retornoDescanso,
      Salida: salida
    };

    if (mapActual[tipo]) {
      showToast(
        `Aviso: "${tipo}" ya fue registrada hoy a las ${mapActual[tipo]?.hora}.`,
        'warning'
      );
      return;
    }

    if (tipo === 'Salida a Descanso' && !entrada) {
      showToast('Bloqueado: Debe registrar primero la "Entrada".', 'error');
      return;
    }
    if (tipo === 'Retorno de Descanso' && !salidaDescanso) {
      showToast('Bloqueado: Debe registrar primero la "Salida a Descanso".', 'error');
      return;
    }
    if (tipo === 'Salida' && !retornoDescanso) {
      showToast('Bloqueado: Debe completar antes el "Retorno de Descanso".', 'error');
      return;
    }

    ejecutarGuardadoMarcaje(tipo);
  };

  // Eliminar un marcaje individual (solo supervisor)
  const eliminarMarcajeIndividual = (idMarcaje: number) => {
    setMarcajes((prev) => prev.filter((m) => m.id !== idMarcaje));
    showToast('Registro de marcaje eliminado.', 'info');
  };

  // Listado en vivo de hoy
  const marcajesHoy = useMemo(() => {
    return marcajes.filter((m) => m.fecha === hoyStr);
  }, [marcajes, hoyStr]);

  const marcajesHoyFiltrados = useMemo(() => {
    const q = cleanStr(liveSearchQuery);
    if (!q) return marcajesHoy;
    return marcajesHoy.filter((item) => {
      return (
        cleanStr(item.locker).includes(q) ||
        cleanStr(item.nombre).includes(q) ||
        cleanStr(item.tipo).includes(q) ||
        cleanStr(item.area).includes(q)
      );
    });
  }, [marcajesHoy, liveSearchQuery]);

  // Matriz Consolidada de Reportes (4 registros por colaborador por fecha)
  const matrizReporte = useMemo(() => {
    const ahora = new Date();

    const marcajesFiltrados = marcajes.filter((m) => {
      const itemDate = new Date(m.timestamp);
      if (filtroPeriodo === 'diario') {
        if (m.fecha !== hoyStr) return false;
      } else if (filtroPeriodo === 'semanal') {
        const diffDays = (ahora.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24);
        if (diffDays > 7) return false;
      } else if (filtroPeriodo === 'quincenal') {
        const diffDays = (ahora.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24);
        if (diffDays > 15) return false;
      } else if (filtroPeriodo === 'mensual') {
        const diffDays = (ahora.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24);
        if (diffDays > 30) return false;
      }
      return true;
    });

    const grupos: Record<
      string,
      {
        key: string;
        fecha: string;
        locker: string;
        nombre: string;
        area: string;
        entrada: string;
        salidaDescanso: string;
        retornoDescanso: string;
        salida: string;
        timestamp: string;
      }
    > = {};

    marcajesFiltrados.forEach((m) => {
      const idClave = m.colabId || m.nombre;
      const key = `${m.fecha}_${idClave}`;
      if (!grupos[key]) {
        grupos[key] = {
          key,
          fecha: m.fecha,
          locker: m.locker && m.locker !== 'S/N' ? m.locker : 'S/N',
          nombre: m.nombre,
          area: m.area,
          entrada: '--:--',
          salidaDescanso: '--:--',
          retornoDescanso: '--:--',
          salida: '--:--',
          timestamp: m.timestamp
        };
      }

      if (m.tipo === 'Entrada') grupos[key].entrada = m.hora;
      else if (m.tipo === 'Salida a Descanso') grupos[key].salidaDescanso = m.hora;
      else if (m.tipo === 'Retorno de Descanso') grupos[key].retornoDescanso = m.hora;
      else if (m.tipo === 'Salida') grupos[key].salida = m.hora;
    });

    let lista = Object.values(grupos);

    if (filtroArea) {
      lista = lista.filter((item) => item.area === filtroArea);
    }

    const q = cleanStr(filtroBusqueda);
    if (q) {
      lista = lista.filter(
        (item) => cleanStr(item.locker).includes(q) || cleanStr(item.nombre).includes(q)
      );
    }

    lista.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
    return lista;
  }, [marcajes, filtroPeriodo, filtroArea, filtroBusqueda, hoyStr]);

  // Cálculos del Módulo de Gráficos de Retardos (> 5 min)
  const datosAnaliticaRetardos = useMemo(() => {
    const ahora = new Date();

    const entradasFiltradas = marcajes.filter((m) => {
      if (m.tipo !== 'Entrada') return false;
      const itemDate = new Date(m.timestamp);
      if (filtroGraficoPeriodo === 'diario') {
        if (m.fecha !== hoyStr) return false;
      } else if (filtroGraficoPeriodo === 'semanal') {
        const diffDays = (ahora.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24);
        if (diffDays > 7) return false;
      } else if (filtroGraficoPeriodo === 'quincenal') {
        const diffDays = (ahora.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24);
        if (diffDays > 15) return false;
      } else if (filtroGraficoPeriodo === 'mensual') {
        const diffDays = (ahora.getTime() - itemDate.getTime()) / (1000 * 60 * 60 * 24);
        if (diffDays > 30) return false;
      }

      if (filtroGraficoArea && m.area !== filtroGraficoArea) return false;
      return true;
    });

    let totalEvaluadas = 0;
    let puntuales = 0;
    const casosRetardoCritico: Array<
      Marcaje & {
        turno: string;
        minutosRetardo: number;
      }
    > = [];

    entradasFiltradas.forEach((item) => {
      const colabRef = personal.find((p) => p.id === item.colabId || p.nombre === item.nombre);
      const res = calcularRetardoEntrada(item.hora, filtroGraficoTurno, colabRef?.turnoBase);
      if (!res) return;

      if (filtroGraficoTurno !== 'todos' && !res.turno.includes(filtroGraficoTurno)) return;

      totalEvaluadas++;
      if (res.esRetardoCritico) {
        casosRetardoCritico.push({
          ...item,
          turno: res.turno,
          minutosRetardo: res.minutosRetardo
        });
      } else {
        puntuales++;
      }
    });

    const sumaMinutos = casosRetardoCritico.reduce((acc, c) => acc + c.minutosRetardo, 0);
    const promMinutos =
      casosRetardoCritico.length > 0 ? Math.round(sumaMinutos / casosRetardoCritico.length) : 0;

    // Agrupar para gráfico de barras por colaborador
    const resumenPorColab: Record<string, number> = {};
    casosRetardoCritico.forEach((c) => {
      const label = c.locker && c.locker !== 'S/N' ? `#${c.locker} ${c.nombre}` : c.nombre;
      resumenPorColab[label] = (resumenPorColab[label] || 0) + c.minutosRetardo;
    });

    const barrasColaboradores = Object.entries(resumenPorColab)
      .map(([label, minutos]) => ({ label, minutos }))
      .sort((a, b) => b.minutos - a.minutos);

    // Conteo por turnos
    const conteoTurnos = {
      '07:00 AM': 0,
      '08:00 AM': 0,
      '09:00 AM': 0
    };
    casosRetardoCritico.forEach((c) => {
      if (c.turno.includes('07:00')) conteoTurnos['07:00 AM']++;
      else if (c.turno.includes('08:00')) conteoTurnos['08:00 AM']++;
      else if (c.turno.includes('09:00')) conteoTurnos['09:00 AM']++;
    });

    return {
      totalEvaluadas,
      puntuales,
      casosRetardoCritico,
      promMinutos,
      barrasColaboradores,
      conteoTurnos
    };
  }, [marcajes, personal, filtroGraficoPeriodo, filtroGraficoTurno, filtroGraficoArea, hoyStr]);

  // Cambio de pestaña con verificación de rol
  const handleSwitchTab = (tab: 'terminal' | 'reportes' | 'config' | 'graficos') => {
    if (perfilActual !== 'supervisor' && tab !== 'terminal') {
      showToast('Módulo protegido. Requiere perfil de Supervisor.', 'warning');
      setClaveSupervisorInput('');
      setModalLoginSupervisorOpen(true);
      return;
    }
    setActiveTab(tab);
  };

  // Login Supervisor
  const verificarClaveSupervisor = (e: React.FormEvent) => {
    e.preventDefault();
    if (claveSupervisorInput.trim() === obtenerClaveSupervisor()) {
      setPerfilActual('supervisor');
      setModalLoginSupervisorOpen(false);
      showToast('¡Sesión de Supervisor iniciada! Módulos administrativos activos.', 'success');
    } else {
      showToast('Clave incorrecta. Acceso denegado.', 'error');
    }
  };

  const cerrarSesionSupervisor = () => {
    setPerfilActual('auxiliar');
    setCustomClockTime('');
    setActiveTab('terminal');
    showToast('Sesión cerrada. Modo Auxiliar activo (solo terminal).', 'info');
  };

  const guardarNuevaClaveSupervisor = (e: React.FormEvent) => {
    e.preventDefault();
    if (claveActualInput.trim() !== obtenerClaveSupervisor()) {
      showToast('La clave actual no es correcta.', 'error');
      return;
    }
    if (nuevaClaveInput.trim().length < 4) {
      showToast('La nueva clave debe tener al menos 4 caracteres.', 'warning');
      return;
    }
    localStorage.setItem(CLAVE_STORAGE_KEY, nuevaClaveInput.trim());
    setModalCambiarClaveOpen(false);
    showToast('Clave de supervisor actualizada correctamente.', 'success');
  };

  // Exportar Excel Reporte
  const exportarExcelReporte = () => {
    if (matrizReporte.length === 0) {
      showToast('No existen datos en el reporte para exportar', 'warning');
      return;
    }

    const formatted = matrizReporte.map((item) => ({
      Fecha: item.fecha,
      'Número de Locker': item.locker,
      Colaborador: item.nombre,
      Área: item.area,
      '1. Entrada': item.entrada,
      '2. Salida a Descanso': item.salidaDescanso,
      '3. Retorno de Descanso': item.retornoDescanso,
      '4. Salida': item.salida
    }));

    const ws = XLSX.utils.json_to_sheet(formatted);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'MangoCenter_Asistencia');
    ws['!cols'] = [
      { wch: 14 },
      { wch: 18 },
      { wch: 32 },
      { wch: 25 },
      { wch: 15 },
      { wch: 20 },
      { wch: 22 },
      { wch: 15 }
    ];

    XLSX.writeFile(
      wb,
      `MangoCenter_Jornadas_${filtroPeriodo.toUpperCase()}_${new Date().toISOString().slice(0, 10)}.xlsx`
    );
    showToast('Reporte en Excel exportado con éxito', 'success');
  };

  // Imprimir Reporte Formal
  const imprimirReporteFormal = () => {
    window.print();
  };

  // Gestión de Personal (Agregar / Editar)
  const abrirModalAgregarPersonal = () => {
    setEditingColab(null);
    setNuevoNombre('');
    setNuevaArea('');
    setNuevoLocker('');
    setNuevoTurnoBase('auto');
    setModalNuevoPersonalOpen(true);
  };

  const abrirModalEditarPersonal = (colab: Colaborador) => {
    setEditingColab(colab);
    setNuevoNombre(colab.nombre);
    setNuevaArea(colab.area);
    setNuevoLocker(colab.locker === 'S/N' ? '' : colab.locker);
    setNuevoTurnoBase(colab.turnoBase || 'auto');
    setModalNuevoPersonalOpen(true);
  };

  const guardarPersonal = (e: React.FormEvent) => {
    e.preventDefault();
    const nombre = nuevoNombre.trim();
    const area = nuevaArea.trim() || 'General';
    const lockerRaw = nuevoLocker.trim().toUpperCase();
    const locker = lockerRaw ? lockerRaw : 'S/N';

    if (!nombre) {
      showToast('El nombre y apellido son obligatorios', 'warning');
      return;
    }

    if (locker !== 'S/N') {
      const existeLocker = personal.find(
        (p) =>
          p.id !== editingColab?.id &&
          p.locker &&
          p.locker !== 'S/N' &&
          cleanStr(p.locker) === cleanStr(locker)
      );
      if (existeLocker) {
        showToast(`El locker "${locker}" ya está asignado a ${existeLocker.nombre}`, 'error');
        return;
      }
    }

    if (editingColab) {
      setPersonal((prev) =>
        prev.map((p) =>
          p.id === editingColab.id
            ? { ...p, nombre, area, locker, turnoBase: nuevoTurnoBase }
            : p
        )
      );
      showToast(`Datos de ${nombre} actualizados correctamente`, 'success');
    } else {
      const nuevoColab: Colaborador = {
        id: 'COLAB_' + Date.now(),
        locker,
        nombre,
        area,
        turnoBase: nuevoTurnoBase
      };
      setPersonal((prev) => [...prev, nuevoColab]);
      showToast(`Colaborador ${nombre} agregado exitosamente`, 'success');
    }

    setModalNuevoPersonalOpen(false);
  };

  const eliminarColaborador = (colab: Colaborador) => {
    setConfirmDialog({
      open: true,
      title: 'Eliminar Colaborador',
      message: `¿Está seguro de retirar a "${colab.nombre}" de la nómina activa?`,
      confirmLabel: 'Sí, eliminar',
      onConfirm: () => {
        setPersonal((prev) => prev.filter((p) => p.id !== colab.id));
        if (colaboradorActual?.id === colab.id) {
          setColaboradorActual(null);
        }
        setConfirmDialog(null);
        showToast(`Colaborador ${colab.nombre} eliminado de la lista`, 'info');
      }
    });
  };

  // Carga masiva Excel
  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result;
        if (!buffer) return;
        const data = new Uint8Array(buffer as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
          defval: ''
        });

        if (rawRows.length === 0) {
          showToast('El archivo Excel está vacío', 'error');
          return;
        }

        const parsedList: Colaborador[] = [];
        rawRows.forEach((row, idx) => {
          const lockerVal =
            row['Número de Locker'] ||
            row['Numero de Locker'] ||
            row['Locker'] ||
            row['LOCKER'] ||
            row['N° Locker'];
          const nombreVal =
            row['Nombre y Apellido'] ||
            row['Nombre'] ||
            row['Colaborador'] ||
            row['Empleado'];
          const areaVal =
            row['Área de Trabajo'] ||
            row['Area de Trabajo'] ||
            row['Área'] ||
            row['Area'] ||
            row['Departamento'];

          if (nombreVal) {
            parsedList.push({
              id: `COLAB_${Date.now()}_${idx}`,
              locker: lockerVal ? String(lockerVal).trim() : 'S/N',
              nombre: String(nombreVal).trim(),
              area: String(areaVal || 'General').trim(),
              turnoBase: 'auto'
            });
          }
        });

        if (parsedList.length === 0) {
          showToast('No se identificó la columna obligatoria (Nombre y Apellido)', 'error');
          return;
        }

        setPersonal(parsedList);
        showToast(`¡Se cargaron ${parsedList.length} colaboradores con éxito!`, 'success');
        e.target.value = '';
      } catch {
        showToast('Error al procesar el archivo Excel', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const descargarPlantillaExcel = () => {
    const plantilla = [
      {
        'Número de Locker': '101',
        'Nombre y Apellido': 'Carlos Eduardo Mendoza',
        'Área de Trabajo': 'Piso de Venta / Ferretería'
      },
      {
        'Número de Locker': '102',
        'Nombre y Apellido': 'Mariana Sofia Silva',
        'Área de Trabajo': 'Cajas y Facturación'
      },
      {
        'Número de Locker': '',
        'Nombre y Apellido': 'José Gregorio Ramos',
        'Área de Trabajo': 'Almacén Central / Despacho'
      }
    ];
    const ws = XLSX.utils.json_to_sheet(plantilla);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Personal_MangoCenter');
    ws['!cols'] = [{ wch: 18 }, { wch: 30 }, { wch: 25 }];
    XLSX.writeFile(wb, 'Plantilla_Personal_MangoCenter.xlsx');
  };

  // Cargar datos demostrativos de jornadas y retardos para pruebas rápidas
  const generarDatosDemostrativos = () => {
    const baseNow = new Date();
    const fechaHoy = baseNow.toLocaleDateString('es-ES');
    const ayerDate = new Date(baseNow.getTime() - 24 * 60 * 60 * 1000);
    const fechaAyer = ayerDate.toLocaleDateString('es-ES');

    const demoMarcajes: Marcaje[] = [
      // Carlos Eduardo Mendoza (#101) - Hoy con retardo > 5 min en turno 07:00 AM (07:18:20)
      {
        id: Date.now() - 50000,
        colabId: '101',
        timestamp: baseNow.toISOString(),
        fecha: fechaHoy,
        hora: '07:18:20',
        locker: '101',
        nombre: 'Carlos Eduardo Mendoza',
        area: 'Piso de Venta / Ferretería',
        tipo: 'Entrada'
      },
      {
        id: Date.now() - 45000,
        colabId: '101',
        timestamp: baseNow.toISOString(),
        fecha: fechaHoy,
        hora: '12:02:10',
        locker: '101',
        nombre: 'Carlos Eduardo Mendoza',
        area: 'Piso de Venta / Ferretería',
        tipo: 'Salida a Descanso'
      },
      {
        id: Date.now() - 40000,
        colabId: '101',
        timestamp: baseNow.toISOString(),
        fecha: fechaHoy,
        hora: '13:00:45',
        locker: '101',
        nombre: 'Carlos Eduardo Mendoza',
        area: 'Piso de Venta / Ferretería',
        tipo: 'Retorno de Descanso'
      },
      // Mariana Sofia Silva (#102) - Hoy puntual turno 08:00 AM (08:03:12)
      {
        id: Date.now() - 35000,
        colabId: '102',
        timestamp: baseNow.toISOString(),
        fecha: fechaHoy,
        hora: '08:03:12',
        locker: '102',
        nombre: 'Mariana Sofia Silva',
        area: 'Cajas y Facturación',
        tipo: 'Entrada'
      },
      {
        id: Date.now() - 30000,
        colabId: '102',
        timestamp: baseNow.toISOString(),
        fecha: fechaHoy,
        hora: '12:30:00',
        locker: '102',
        nombre: 'Mariana Sofia Silva',
        area: 'Cajas y Facturación',
        tipo: 'Salida a Descanso'
      },
      // José Gregorio Ramos (#103) - Hoy con retardo > 5 min en turno 07:00 AM (07:24:10)
      {
        id: Date.now() - 28000,
        colabId: '103',
        timestamp: baseNow.toISOString(),
        fecha: fechaHoy,
        hora: '07:24:10',
        locker: '103',
        nombre: 'José Gregorio Ramos',
        area: 'Almacén Central / Despacho',
        tipo: 'Entrada'
      },
      // Elena Beatriz Rivas (#104) - Hoy con retardo > 5 min en turno 09:00 AM (09:14:05)
      {
        id: Date.now() - 25000,
        colabId: '104',
        timestamp: baseNow.toISOString(),
        fecha: fechaHoy,
        hora: '09:14:05',
        locker: '104',
        nombre: 'Elena Beatriz Rivas',
        area: 'Atención al Cliente',
        tipo: 'Entrada'
      },
      // Jornada completa de ayer para Mariana (#102) con retardo > 5 min en turno 08:00 AM (08:16:40)
      {
        id: Date.now() - 90000,
        colabId: '102',
        timestamp: ayerDate.toISOString(),
        fecha: fechaAyer,
        hora: '08:16:40',
        locker: '102',
        nombre: 'Mariana Sofia Silva',
        area: 'Cajas y Facturación',
        tipo: 'Entrada'
      },
      {
        id: Date.now() - 89000,
        colabId: '102',
        timestamp: ayerDate.toISOString(),
        fecha: fechaAyer,
        hora: '12:15:00',
        locker: '102',
        nombre: 'Mariana Sofia Silva',
        area: 'Cajas y Facturación',
        tipo: 'Salida a Descanso'
      },
      {
        id: Date.now() - 88000,
        colabId: '102',
        timestamp: ayerDate.toISOString(),
        fecha: fechaAyer,
        hora: '13:14:22',
        locker: '102',
        nombre: 'Mariana Sofia Silva',
        area: 'Cajas y Facturación',
        tipo: 'Retorno de Descanso'
      },
      {
        id: Date.now() - 87000,
        colabId: '102',
        timestamp: ayerDate.toISOString(),
        fecha: fechaAyer,
        hora: '17:05:10',
        locker: '102',
        nombre: 'Mariana Sofia Silva',
        area: 'Cajas y Facturación',
        tipo: 'Salida'
      }
    ];

    setMarcajes(demoMarcajes);
    showToast('Se cargaron jornadas y casos de retardo de ejemplo.', 'success');
  };

  const badgeColors: Record<TipoMarcaje, string> = {
    Entrada: 'bg-emerald-50 text-emerald-800 border-emerald-300',
    'Salida a Descanso': 'bg-amber-50 text-amber-800 border-amber-300',
    'Retorno de Descanso': 'bg-blue-50 text-mango-blue border-blue-300',
    Salida: 'bg-rose-50 text-mango-red border-rose-300'
  };

  // Valores para el gráfico de dona SVG
  const totalTurnosRetardo =
    datosAnaliticaRetardos.conteoTurnos['07:00 AM'] +
    datosAnaliticaRetardos.conteoTurnos['08:00 AM'] +
    datosAnaliticaRetardos.conteoTurnos['09:00 AM'];

  return (
    <div className="bg-slate-100 text-slate-800 min-h-screen flex flex-col font-sans antialiased selection:bg-mango-blue selection:text-white">
      {/* HEADER OFICIAL MANGO CENTER */}
      <header className="bg-mango-dark text-white shadow-lg border-b-4 border-mango-red no-print sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-2 sm:py-2.5">
          {/* Fila Superior: Logo, Identidad, Reloj y Perfil */}
          <div className="flex items-center justify-between gap-2 sm:gap-4">
            <div className="flex items-center gap-2 sm:gap-3.5">
              <div className="bg-black/80 p-1 sm:p-1.5 px-2 sm:px-3 rounded-xl border border-white/10 flex items-center shadow-inner">
                <img
                  src="/assets/aistudio/mango_logo.jpg"
                  alt="Mango Center Hiperferretería"
                  className="h-8 sm:h-11 w-auto object-contain"
                />
              </div>
              <div>
                <h1 className="text-xs sm:text-sm font-extrabold uppercase tracking-wider text-white leading-tight">
                  Mango Center
                </h1>
                <p className="text-[10px] sm:text-[11px] text-slate-300 font-medium hidden sm:block">
                  Control de Asistencia y Jornada
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 sm:gap-3">
              {/* Reloj Digital Corporativo */}
              <div className="text-right bg-slate-900/90 px-2.5 sm:px-3.5 py-1 sm:py-1.5 rounded-lg border border-slate-700/60 shadow-sm">
                <div className="text-xs sm:text-base md:text-lg font-mono font-bold tracking-wider text-sky-400 tabular-nums">
                  {timeStr}
                </div>
                <div className="text-[9px] sm:text-[10px] md:text-xs text-slate-300 capitalize font-medium hidden sm:block">
                  {dateStr}
                </div>
              </div>

              {/* Control de Perfiles (Auxiliar / Supervisor) */}
              <div className="flex items-center gap-1.5 sm:gap-2">
                {perfilActual === 'auxiliar' ? (
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="inline-flex items-center gap-1 px-2 py-1 rounded-md text-[10px] sm:text-xs font-semibold bg-slate-800 text-slate-300 border border-slate-700">
                      <UserCheck className="w-3.5 h-3.5 text-sky-400" />
                      <span className="hidden sm:inline">Auxiliar</span>
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        setClaveSupervisorInput('');
                        setModalLoginSupervisorOpen(true);
                      }}
                      className="btn-mango-blue inline-flex items-center gap-1.5 px-2.5 py-1.5 sm:px-3 sm:py-1.5 text-white text-[11px] sm:text-xs font-bold rounded-lg shadow-sm border border-blue-400/30 transition active:scale-[0.98] whitespace-nowrap cursor-pointer"
                      style={{ backgroundColor: '#0d5ca4', color: '#ffffff' }}
                    >
                      <Lock className="w-3.5 h-3.5 text-amber-300" /> Supervisor
                    </button>
                  </div>
                ) : (
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    <span className="inline-flex items-center gap-1 px-2 py-1 rounded-md text-[10px] sm:text-xs font-bold bg-amber-500/20 text-amber-300 border border-amber-400/40">
                      <ShieldCheck className="w-3.5 h-3.5 text-amber-400" />
                      <span className="hidden md:inline">Supervisor</span>
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        setClaveActualInput('');
                        setNuevaClaveInput('');
                        setModalCambiarClaveOpen(true);
                      }}
                      title="Cambiar clave de acceso"
                      className="p-1.5 sm:px-2.5 sm:py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold rounded-lg border border-slate-600 transition cursor-pointer"
                    >
                      <Key className="w-3.5 h-3.5 text-sky-400" />
                    </button>
                    <button
                      type="button"
                      onClick={cerrarSesionSupervisor}
                      title="Cerrar sesión de supervisor"
                      className="btn-mango-red inline-flex items-center gap-1 px-2.5 py-1.5 sm:px-3 sm:py-1.5 text-white text-[11px] sm:text-xs font-bold rounded-lg shadow-sm transition active:scale-[0.98] whitespace-nowrap cursor-pointer"
                      style={{ backgroundColor: '#e5242d', color: '#ffffff' }}
                    >
                      <LogOut className="w-3.5 h-3.5" />
                      <span className="hidden sm:inline">Salir</span>
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Fila Inferior: Navegación de Pestañas */}
          <div className="mt-2 pt-2 border-t border-slate-700/60 flex items-center justify-between gap-2 overflow-x-auto no-scrollbar">
            <nav className="flex items-center gap-1.5 min-w-max">
              <button
                type="button"
                onClick={() => handleSwitchTab('terminal')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm transition whitespace-nowrap cursor-pointer ${
                  activeTab === 'terminal'
                    ? 'btn-mango-blue font-bold text-white shadow-sm'
                    : 'font-medium text-slate-300 hover:text-white hover:bg-white/10'
                }`}
                style={activeTab === 'terminal' ? { backgroundColor: '#0d5ca4', color: '#ffffff' } : {}}
              >
                <Fingerprint className="w-4 h-4" /> Terminal
              </button>

              {perfilActual === 'supervisor' && (
                <>
                  <button
                    type="button"
                    onClick={() => handleSwitchTab('reportes')}
                    className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm transition whitespace-nowrap cursor-pointer ${
                      activeTab === 'reportes'
                        ? 'btn-mango-blue font-bold text-white shadow-sm'
                        : 'font-medium text-slate-300 hover:text-white hover:bg-white/10'
                    }`}
                    style={activeTab === 'reportes' ? { backgroundColor: '#0d5ca4', color: '#ffffff' } : {}}
                  >
                    <FileSpreadsheet className="w-4 h-4" /> Reportes
                  </button>
                  <button
                    type="button"
                    onClick={() => handleSwitchTab('config')}
                    className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm transition whitespace-nowrap cursor-pointer ${
                      activeTab === 'config'
                        ? 'btn-mango-blue font-bold text-white shadow-sm'
                        : 'font-medium text-slate-300 hover:text-white hover:bg-white/10'
                    }`}
                    style={activeTab === 'config' ? { backgroundColor: '#0d5ca4', color: '#ffffff' } : {}}
                  >
                    <Database className="w-4 h-4" /> Personal ({personal.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => handleSwitchTab('graficos')}
                    className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm transition whitespace-nowrap cursor-pointer ${
                      activeTab === 'graficos'
                        ? 'btn-mango-blue font-bold text-white shadow-sm'
                        : 'font-medium text-slate-300 hover:text-white hover:bg-white/10'
                    }`}
                    style={activeTab === 'graficos' ? { backgroundColor: '#0d5ca4', color: '#ffffff' } : {}}
                  >
                    <BarChart3 className="w-4 h-4 text-mango-red" /> Gráficos de Retardos
                  </button>
                </>
              )}
            </nav>

            {perfilActual === 'supervisor' && (
              <button
                type="button"
                onClick={generarDatosDemostrativos}
                className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-semibold bg-slate-800 hover:bg-slate-700 text-sky-300 border border-slate-700 transition whitespace-nowrap shrink-0 cursor-pointer"
                title="Cargar jornadas y retardos de prueba para visualizar reportes y gráficos"
              >
                <Sparkles className="w-3.5 h-3.5 text-amber-400" /> Cargar Datos Demo
              </button>
            )}
          </div>
        </div>
      </header>

      {/* CONTENEDOR PRINCIPAL */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-3 sm:p-6 lg:p-8 no-print">
        {/* Notificaciones Emergentes (Toast) */}
        <div className="fixed top-4 right-3 sm:top-5 sm:right-5 z-50 flex flex-col gap-2 pointer-events-none max-w-[90vw] sm:max-w-md">
          {toasts.map((t) => {
            const colors = {
              success: 'bg-emerald-700 text-white border-emerald-800',
              error: 'bg-mango-red text-white border-mango-redDark',
              warning: 'bg-amber-600 text-white border-amber-700',
              info: 'bg-mango-blue text-white border-mango-blueDark'
            };
            return (
              <div
                key={t.id}
                className={`pointer-events-auto px-3.5 py-2.5 sm:px-4 sm:py-3 rounded-xl shadow-xl border text-xs font-bold flex items-center gap-2 transition-all duration-200 ${colors[t.type]}`}
              >
                <span>{t.msg}</span>
              </div>
            );
          })}
        </div>

        {/* VISTA 1: TERMINAL DE REGISTRO */}
        {activeTab === 'terminal' && (
          <section className="space-y-4 sm:space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 sm:gap-6 items-start">
              {/* Tarjeta de Búsqueda e Identificación */}
              <div className="lg:col-span-7 bg-white rounded-2xl shadow-sm border border-slate-200/90 p-4 sm:p-6 space-y-4 sm:space-y-6">
                <div className="border-b border-slate-100 pb-3 sm:pb-4 flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <h2 className="text-sm sm:text-lg font-bold text-mango-blueDark flex items-center gap-1.5 sm:gap-2">
                      <Search className="w-4 h-4 sm:w-5 sm:h-5 text-mango-red" /> Búsqueda de Colaborador
                    </h2>
                    <p className="text-[11px] sm:text-xs text-slate-500 mt-0.5">
                      Por <strong>Nombre y Apellido</strong> o <strong>N° de Locker</strong>.
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    {perfilActual === 'supervisor' && (
                      <div className="flex items-center gap-1.5 bg-amber-50 border border-amber-200 px-2.5 py-1 rounded-lg">
                        <Clock className="w-3.5 h-3.5 text-amber-700" />
                        <span className="text-[10px] font-bold text-amber-800">Hora Test:</span>
                        <input
                          type="time"
                          value={customClockTime}
                          onChange={(e) => setCustomClockTime(e.target.value)}
                          className="text-xs font-mono font-bold text-amber-900 bg-transparent outline-none"
                          title="Opcional (Solo Supervisor): Simular hora de marcaje para pruebas de retardo"
                        />
                        {customClockTime && (
                          <button
                            type="button"
                            onClick={() => setCustomClockTime('')}
                            className="text-[10px] text-mango-red font-bold hover:underline"
                          >
                            Real
                          </button>
                        )}
                      </div>
                    )}
                    <span className="text-[11px] font-semibold text-mango-blue">
                      Terminal Operativo
                    </span>
                  </div>
                </div>

                {/* Buscador con Autocompletado */}
                <form onSubmit={handleLockerSearch} className="relative">
                  <div className="flex gap-2">
                    <div className="relative flex-1">
                      <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                        <Key className="w-4 h-4 sm:w-5 sm:h-5 text-mango-blue" />
                      </div>
                      <input
                        ref={searchInputRef}
                        type="text"
                        value={lockerQuery}
                        onChange={(e) => {
                          setLockerQuery(e.target.value);
                          setShowSuggestions(true);
                        }}
                        onFocus={() => {
                          if (lockerQuery.trim()) setShowSuggestions(true);
                        }}
                        placeholder="Nombre, Apellido o Locker (ej. 101)..."
                        autoComplete="off"
                        className="w-full pl-9 sm:pl-11 pr-8 py-2.5 sm:py-3 text-base sm:text-lg font-semibold bg-slate-50 border-2 border-slate-300 rounded-xl focus:bg-white focus:border-mango-blue focus:outline-none focus:ring-2 focus:ring-mango-blue/20 transition tracking-wide"
                      />
                      {lockerQuery && (
                        <button
                          type="button"
                          onClick={() => {
                            setLockerQuery('');
                            setColaboradorActual(null);
                            setShowSuggestions(false);
                            searchInputRef.current?.focus();
                          }}
                          className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600 cursor-pointer"
                        >
                          <X className="w-4 h-4" />
                        </button>
                      )}

                      {/* Sugerencias flotantes */}
                      {showSuggestions && lockerQuery.trim().length > 0 && (
                        <div
                          ref={suggestionsBoxRef}
                          className="absolute left-0 right-0 top-full mt-1 bg-white border border-slate-200 rounded-xl shadow-xl z-50 max-h-60 overflow-y-auto divide-y divide-slate-100"
                        >
                          {suggestions.length === 0 ? (
                            <div className="p-3 text-xs text-slate-400 text-center">
                              No se encontraron colaboradores con "{lockerQuery}"
                            </div>
                          ) : (
                            suggestions.map((m) => (
                              <button
                                key={m.id}
                                type="button"
                                onClick={() => seleccionarColaborador(m)}
                                className="w-full text-left p-2.5 px-3 flex items-center justify-between hover:bg-blue-50/70 cursor-pointer transition"
                              >
                                <div className="flex items-center gap-2.5">
                                  <span
                                    className={`font-mono font-bold text-xs px-1.5 py-0.5 rounded tabular-nums ${
                                      m.locker && m.locker !== 'S/N'
                                        ? 'bg-blue-100 text-mango-blue'
                                        : 'bg-slate-200 text-slate-600'
                                    }`}
                                  >
                                    {m.locker && m.locker !== 'S/N' ? `#${m.locker}` : 'S/N'}
                                  </span>
                                  <div>
                                    <p className="text-xs font-bold text-slate-800 leading-tight">
                                      {m.nombre}
                                    </p>
                                    <p className="text-[10px] text-slate-500">
                                      {m.area}
                                      {m.turnoBase && m.turnoBase !== 'auto'
                                        ? ` · Turno ${m.turnoBase} AM`
                                        : ''}
                                    </p>
                                  </div>
                                </div>
                                <span className="text-[10px] font-semibold text-mango-blue whitespace-nowrap">
                                  Seleccionar &rarr;
                                </span>
                              </button>
                            ))
                          )}
                        </div>
                      )}
                    </div>
                    <button
                      type="submit"
                      className="btn-mango-blue px-3.5 sm:px-5 py-2.5 sm:py-3 text-white font-bold text-xs sm:text-sm rounded-xl flex items-center gap-1.5 shadow-md shadow-mango-blue/20 transition active:scale-[0.98] whitespace-nowrap cursor-pointer"
                      style={{ backgroundColor: '#0d5ca4', color: '#ffffff' }}
                    >
                      <Search className="w-4 h-4 sm:w-5 sm:h-5" />
                      <span className="hidden sm:inline">Consultar</span>
                    </button>
                  </div>
                </form>

                {/* Accesos rápidos de colaboradores frecuentes */}
                <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar pt-0.5">
                  <span className="text-[11px] font-semibold text-slate-400 shrink-0 mr-1">
                    Acceso rápido:
                  </span>
                  {personal.slice(0, 5).map((p) => (
                    <button
                      key={p.id}
                      type="button"
                      onClick={() => seleccionarColaborador(p)}
                      className={`px-2.5 py-1 rounded-lg text-[11px] font-semibold border transition whitespace-nowrap shrink-0 cursor-pointer ${
                        colaboradorActual?.id === p.id
                          ? 'bg-mango-blue text-white border-mango-blue'
                          : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'
                      }`}
                    >
                      {p.locker && p.locker !== 'S/N' ? `#${p.locker} ` : ''}
                      {p.nombre.split(' ').slice(0, 2).join(' ')}
                    </button>
                  ))}
                </div>

                {/* Ficha del Colaborador y Botones de Marcaje */}
                <div className="bg-slate-50 border border-slate-200 rounded-2xl p-3.5 sm:p-5 transition-all">
                  {!colaboradorActual ? (
                    <div className="space-y-4">
                      <div className="py-4 sm:py-6 text-center text-slate-400 border-b border-slate-200/80">
                        <div
                          className="w-12 h-12 sm:w-14 sm:h-14 mx-auto mb-2 bg-white rounded-full flex items-center justify-center border border-slate-200 shadow-sm"
                          style={{ backgroundColor: '#ffffff' }}
                        >
                          <UserCheck className="w-6 h-6 sm:w-8 sm:h-8" style={{ color: '#0d5ca4' }} />
                        </div>
                        <p className="text-xs sm:text-sm font-bold text-slate-700">
                          Busque o seleccione un colaborador para registrar marcaje
                        </p>
                        <p className="text-[10px] sm:text-xs text-slate-500 mt-0.5">
                          Escriba el nombre, apellido o N° de locker arriba o use los accesos rápidos
                        </p>
                      </div>

                      {/* Botones de Marcaje (Siempre Visibles en Estado de Espera) */}
                      <div>
                        <div className="flex items-center justify-between mb-2">
                          <label className="text-[10px] sm:text-xs font-bold text-slate-600 uppercase tracking-wider">
                            Botones de Marcaje (Esperando Selección):
                          </label>
                          <span
                            className="text-[10px] font-semibold px-2 py-0.5 rounded-full"
                            style={{ backgroundColor: '#dbeafe', color: '#0d5ca4' }}
                          >
                            Modo Terminal
                          </span>
                        </div>
                        <div className="grid grid-cols-2 gap-2.5 sm:gap-3">
                          <button
                            type="button"
                            disabled
                            className="btn-entrada min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm shadow-sm transition opacity-60 cursor-not-allowed"
                            style={{ backgroundColor: '#059669', color: '#ffffff' }}
                            title="Seleccione un colaborador para habilitar Entrada"
                          >
                            <LogIn className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">1. Entrada</span>
                          </button>

                          <button
                            type="button"
                            disabled
                            className="btn-sal-descanso min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm shadow-sm transition opacity-60 cursor-not-allowed"
                            style={{ backgroundColor: '#d97706', color: '#ffffff' }}
                            title="Seleccione un colaborador para habilitar Salida a Descanso"
                          >
                            <Coffee className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">2. Sal. Descanso</span>
                          </button>

                          <button
                            type="button"
                            disabled
                            className="btn-ret-descanso min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm shadow-sm transition opacity-60 cursor-not-allowed"
                            style={{ backgroundColor: '#0d5ca4', color: '#ffffff' }}
                            title="Seleccione un colaborador para habilitar Retorno de Descanso"
                          >
                            <Utensils className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">3. Ret. Descanso</span>
                          </button>

                          <button
                            type="button"
                            disabled
                            className="btn-salida min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm shadow-sm transition opacity-60 cursor-not-allowed"
                            style={{ backgroundColor: '#e5242d', color: '#ffffff' }}
                            title="Seleccione un colaborador para habilitar Salida Final"
                          >
                            <LogOut className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">4. Salida Final</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      <div className="flex items-center justify-between pb-3 border-b border-slate-200/80 gap-2">
                        <div className="flex items-center gap-2.5 sm:gap-3.5">
                          <div
                            className="w-10 h-10 sm:w-12 sm:h-12 rounded-xl text-white font-extrabold flex items-center justify-center text-sm sm:text-lg shadow-sm shrink-0"
                            style={{ backgroundColor: '#0d5ca4', color: '#ffffff' }}
                          >
                            {colaboradorActual.nombre
                              .split(' ')
                              .map((n) => n[0])
                              .filter(Boolean)
                              .slice(0, 2)
                              .join('')}
                          </div>
                          <div>
                            <h3 className="text-sm sm:text-base font-extrabold text-slate-900 leading-tight">
                              {colaboradorActual.nombre}
                            </h3>
                            <div className="flex flex-wrap items-center gap-2 mt-1 text-xs text-slate-600">
                              <span className="font-semibold" style={{ color: '#0d5ca4' }}>
                                {colaboradorActual.area}
                              </span>
                              <span aria-hidden="true">&bull;</span>
                              <span>
                                Locker:{' '}
                                <strong className="font-mono font-bold" style={{ color: '#e5242d' }}>
                                  {colaboradorActual.locker && colaboradorActual.locker !== 'S/N'
                                    ? `#${colaboradorActual.locker}`
                                    : 'S/N'}
                                </strong>
                              </span>
                            </div>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 shrink-0">
                          <div className="text-right">
                            <span className="text-[10px] uppercase font-bold text-slate-400 block tracking-wider">
                              Progreso Hoy
                            </span>
                            <span className="text-xs font-bold text-slate-700 tabular-nums">
                              {estadoMarcajesColaborador.total} de 4 marcajes
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => {
                              setColaboradorActual(null);
                              setLockerQuery('');
                              setAlertaMarcajeAccidental(null);
                              searchInputRef.current?.focus();
                            }}
                            title="Cerrar ficha de colaborador"
                            className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-200/70 rounded-lg transition cursor-pointer"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      </div>

                      {/* Resumen de los 4 registros de hoy */}
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-white p-2.5 rounded-xl border border-slate-200/90 text-center">
                        <div className="p-2 rounded-lg bg-slate-50 border border-slate-100">
                          <div className="text-[10px] font-bold text-slate-500 uppercase truncate">
                            1. Entrada
                          </div>
                          <div
                            className={`text-xs font-mono font-bold mt-0.5 tabular-nums ${
                              estadoMarcajesColaborador.entrada
                                ? 'text-emerald-700'
                                : 'text-slate-400'
                            }`}
                            style={estadoMarcajesColaborador.entrada ? { color: '#047857' } : {}}
                          >
                            {estadoMarcajesColaborador.entrada?.hora || '--:--'}
                          </div>
                        </div>
                        <div className="p-2 rounded-lg bg-slate-50 border border-slate-100">
                          <div className="text-[10px] font-bold text-slate-500 uppercase truncate">
                            2. Sal. Desc.
                          </div>
                          <div
                            className={`text-xs font-mono font-bold mt-0.5 tabular-nums ${
                              estadoMarcajesColaborador.salidaDescanso
                                ? 'text-amber-700'
                                : 'text-slate-400'
                            }`}
                            style={estadoMarcajesColaborador.salidaDescanso ? { color: '#b45309' } : {}}
                          >
                            {estadoMarcajesColaborador.salidaDescanso?.hora || '--:--'}
                          </div>
                        </div>
                        <div className="p-2 rounded-lg bg-slate-50 border border-slate-100">
                          <div className="text-[10px] font-bold text-slate-500 uppercase truncate">
                            3. Ret. Desc.
                          </div>
                          <div
                            className={`text-xs font-mono font-bold mt-0.5 tabular-nums ${
                              estadoMarcajesColaborador.retornoDescanso
                                ? 'text-mango-blue'
                                : 'text-slate-400'
                            }`}
                            style={estadoMarcajesColaborador.retornoDescanso ? { color: '#0d5ca4' } : {}}
                          >
                            {estadoMarcajesColaborador.retornoDescanso?.hora || '--:--'}
                          </div>
                        </div>
                        <div className="p-2 rounded-lg bg-slate-50 border border-slate-100">
                          <div className="text-[10px] font-bold text-slate-500 uppercase truncate">
                            4. Salida
                          </div>
                          <div
                            className={`text-xs font-mono font-bold mt-0.5 tabular-nums ${
                              estadoMarcajesColaborador.salida
                                ? 'text-mango-red'
                                : 'text-slate-400'
                            }`}
                            style={estadoMarcajesColaborador.salida ? { color: '#e5242d' } : {}}
                          >
                            {estadoMarcajesColaborador.salida?.hora || '--:--'}
                          </div>
                        </div>
                      </div>

                      {/* Botones de Marcaje con Secuencia Obligatoria y Estilos HEX Explícitos */}
                      <div>
                        <label className="block text-[10px] sm:text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">
                          Acción a Registrar (Secuencia de Jornada):
                        </label>
                        <div className="grid grid-cols-2 gap-2.5 sm:gap-3">
                          <button
                            type="button"
                            disabled={!estadoMarcajesColaborador.puedeEntrada}
                            onClick={() => registrarMarcaje('Entrada')}
                            className={`btn-entrada min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm text-white shadow-sm transition cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed ${
                              estadoMarcajesColaborador.entrada ? 'line-through' : ''
                            }`}
                            style={{
                              backgroundColor: '#059669',
                              color: '#ffffff',
                              opacity: estadoMarcajesColaborador.puedeEntrada ? 1 : 0.4
                            }}
                          >
                            <LogIn className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">1. Entrada</span>
                          </button>

                          <button
                            type="button"
                            disabled={!estadoMarcajesColaborador.puedeSalDescanso}
                            onClick={() => registrarMarcaje('Salida a Descanso')}
                            className={`btn-sal-descanso min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm text-white shadow-sm transition cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed ${
                              estadoMarcajesColaborador.salidaDescanso ? 'line-through' : ''
                            }`}
                            style={{
                              backgroundColor: '#d97706',
                              color: '#ffffff',
                              opacity: estadoMarcajesColaborador.puedeSalDescanso ? 1 : 0.4
                            }}
                          >
                            <Coffee className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">2. Sal. Descanso</span>
                          </button>

                          <button
                            type="button"
                            disabled={!estadoMarcajesColaborador.puedeRetDescanso}
                            onClick={() => registrarMarcaje('Retorno de Descanso')}
                            className={`btn-ret-descanso min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm text-white shadow-sm transition cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed ${
                              estadoMarcajesColaborador.retornoDescanso ? 'line-through' : ''
                            }`}
                            style={{
                              backgroundColor: '#0d5ca4',
                              color: '#ffffff',
                              opacity: estadoMarcajesColaborador.puedeRetDescanso ? 1 : 0.4
                            }}
                          >
                            <Utensils className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">3. Ret. Descanso</span>
                          </button>

                          <button
                            type="button"
                            disabled={!estadoMarcajesColaborador.puedeSalida}
                            onClick={() => registrarMarcaje('Salida')}
                            className={`btn-salida min-h-[48px] flex items-center justify-center gap-2 py-2.5 px-3 rounded-xl font-bold text-xs sm:text-sm text-white shadow-sm transition cursor-pointer disabled:opacity-40 disabled:cursor-not-allowed ${
                              estadoMarcajesColaborador.salida ? 'line-through' : ''
                            }`}
                            style={{
                              backgroundColor: '#e5242d',
                              color: '#ffffff',
                              opacity: estadoMarcajesColaborador.puedeSalida ? 1 : 0.4
                            }}
                          >
                            <LogOut className="w-4 h-4 sm:w-5 sm:h-5 shrink-0" />
                            <span className="truncate">4. Salida Final</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  )}
                </div>

              </div>

              {/* Historial en Vivo del Día */}
              <div className="lg:col-span-5 bg-white rounded-2xl shadow-sm border border-slate-200/90 p-4 sm:p-6 flex flex-col h-[480px] sm:h-[580px]">
                <div className="flex items-center justify-between pb-3 border-b border-slate-100 mb-3">
                  <div>
                    <h3 className="font-bold text-mango-blueDark flex items-center gap-1.5 text-xs sm:text-sm">
                      <History className="w-4 h-4 text-mango-red" /> Marcajes Recientes (Hoy)
                    </h3>
                    <span className="text-[11px] text-slate-500 font-medium tabular-nums">
                      {marcajesHoy.length} marcaje(s) registrados hoy
                    </span>
                  </div>
                  {marcajesHoy.length > 0 && (
                    <button
                      type="button"
                      onClick={() =>
                        setConfirmDialog({
                          open: true,
                          title: 'Restablecer Marcajes de Hoy',
                          message:
                            '¿Desea limpiar los marcajes registrados exclusivamente para el día de hoy?',
                          confirmLabel: 'Sí, limpiar hoy',
                          onConfirm: () => {
                            setMarcajes((prev) => prev.filter((m) => m.fecha !== hoyStr));
                            setConfirmDialog(null);
                            showToast('Marcajes de hoy restablecidos', 'warning');
                          }
                        })
                      }
                      className="text-xs text-slate-400 hover:text-mango-red transition p-1.5 rounded-lg hover:bg-slate-100 cursor-pointer"
                      title="Restablecer marcajes de hoy"
                    >
                      <RotateCcw className="w-4 h-4" />
                    </button>
                  )}
                </div>

                {/* Buscador del listado en vivo */}
                <div className="mb-3 relative">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                    <Search className="w-3.5 h-3.5 text-slate-400" />
                  </div>
                  <input
                    type="text"
                    value={liveSearchQuery}
                    onChange={(e) => setLiveSearchQuery(e.target.value)}
                    placeholder="Filtrar marcajes de hoy por nombre, locker o tipo..."
                    className="w-full pl-8 pr-3 py-2 text-xs bg-slate-50 border border-slate-200 rounded-lg focus:bg-white focus:border-mango-blue focus:outline-none focus:ring-2 focus:ring-mango-blue/20 transition"
                  />
                </div>

                <div className="flex-1 overflow-y-auto space-y-2 pr-1">
                  {marcajesHoyFiltrados.length === 0 ? (
                    <div className="text-center py-14 text-slate-400">
                      <Clock className="w-8 h-8 mx-auto mb-2 text-slate-300" />
                      <p className="text-xs font-medium">
                        {liveSearchQuery
                          ? 'No se encontraron marcajes con ese filtro.'
                          : 'Aún no se han registrado marcajes en la jornada de hoy.'}
                      </p>
                    </div>
                  ) : (
                    marcajesHoyFiltrados.map((item) => {
                      const evalRetardo =
                        item.tipo === 'Entrada' ? calcularRetardoEntrada(item.hora) : null;
                      return (
                        <div
                          key={item.id}
                          onClick={() => {
                            const found = personal.find(
                              (p) => p.id === item.colabId || p.nombre === item.nombre
                            );
                            if (found) seleccionarColaborador(found);
                          }}
                          className="flex items-center justify-between p-2.5 rounded-xl border border-slate-200 bg-white hover:border-mango-blue/40 transition text-xs cursor-pointer"
                        >
                          <div className="flex items-center gap-2.5 min-w-0">
                            <div className="font-mono font-bold bg-slate-100 px-2 py-1 rounded border border-slate-200 text-mango-blueDark text-[11px] tabular-nums shrink-0">
                              {item.locker && item.locker !== 'S/N' ? `#${item.locker}` : 'S/N'}
                            </div>
                            <div className="min-w-0">
                              <p className="font-bold text-slate-900 leading-tight truncate">
                                {item.nombre}
                              </p>
                              <p className="text-[10px] text-slate-500 truncate">{item.area}</p>
                            </div>
                          </div>
                          <div className="text-right flex items-center gap-2 shrink-0">
                            <div>
                              <span
                                className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold border ${
                                  badgeColors[item.tipo]
                                }`}
                              >
                                {item.tipo}
                              </span>
                              <div className="font-mono text-[11px] text-slate-600 font-bold mt-0.5 tabular-nums flex items-center justify-end gap-1">
                                <span>{item.hora}</span>
                                {evalRetardo?.esRetardoCritico && (
                                  <span className="text-[9px] text-mango-red font-extrabold">
                                    (+{evalRetardo.minutosRetardo}m)
                                  </span>
                                )}
                              </div>
                            </div>
                            {perfilActual === 'supervisor' && (
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  eliminarMarcajeIndividual(item.id);
                                }}
                                title="Eliminar este marcaje"
                                className="p-1 text-slate-300 hover:text-mango-red transition"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>
            </div>
          </section>
        )}

        {/* VISTA 2: REPORTES CONSOLIDADOS POR EMPLEADO (EXCLUSIVO SUPERVISOR) */}
        {activeTab === 'reportes' && perfilActual === 'supervisor' && (
          <section className="space-y-4 sm:space-y-6">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200/90 p-4 sm:p-6 space-y-4 sm:space-y-6">
              {/* Barra superior de Reportes */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
                <div className="flex items-center gap-3">
                  <div className="w-9 h-9 sm:w-10 sm:h-10 rounded-xl bg-mango-blue/10 flex items-center justify-center text-mango-blue border border-mango-blue/20 shrink-0">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm sm:text-lg font-bold text-slate-900 leading-tight">
                      Matriz Consolidada de Jornadas (4 Registros)
                    </h2>
                    <p className="text-[11px] sm:text-xs text-slate-500">
                      Entrada, Descansos y Salida por colaborador.
                    </p>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={exportarExcelReporte}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 sm:px-3.5 sm:py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-xl text-xs font-bold shadow-sm transition active:scale-[0.98] cursor-pointer"
                  >
                    <FileSpreadsheet className="w-4 h-4" /> Exportar a Excel
                  </button>
                  <button
                    type="button"
                    onClick={imprimirReporteFormal}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 sm:px-3.5 sm:py-2 bg-mango-blue hover:bg-mango-blueDark text-white rounded-xl text-xs font-bold shadow-sm transition active:scale-[0.98] cursor-pointer"
                  >
                    <Printer className="w-4 h-4" /> Imprimir / PDF
                  </button>
                </div>
              </div>

              {/* Filtros de Reportes */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-slate-50 p-3 sm:p-4 rounded-xl border border-slate-200 text-xs">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Período de Análisis:
                  </label>
                  <select
                    value={filtroPeriodo}
                    onChange={(e) =>
                      setFiltroPeriodo(
                        e.target.value as 'diario' | 'semanal' | 'quincenal' | 'mensual' | 'todos'
                      )
                    }
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 font-medium focus:ring-2 focus:ring-mango-blue/20 focus:border-mango-blue outline-none"
                  >
                    <option value="diario">Diario (Hoy)</option>
                    <option value="semanal">Semanal (Últimos 7 días)</option>
                    <option value="quincenal">Quincenal (Últimos 15 días)</option>
                    <option value="mensual">Mensual (Últimos 30 días)</option>
                    <option value="todos">Histórico Completo</option>
                  </select>
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Filtrar por Área / Departamento:
                  </label>
                  <select
                    value={filtroArea}
                    onChange={(e) => setFiltroArea(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 font-medium focus:ring-2 focus:ring-mango-blue/20 focus:border-mango-blue outline-none"
                  >
                    <option value="">Todas las áreas</option>
                    {areasDisponibles.map((a) => (
                      <option key={a} value={a}>
                        {a}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Buscar por Locker o Nombre:
                  </label>
                  <input
                    type="text"
                    value={filtroBusqueda}
                    onChange={(e) => setFiltroBusqueda(e.target.value)}
                    placeholder="Ej. 101, Carlos Mendoza..."
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 focus:ring-2 focus:ring-mango-blue/20 focus:border-mango-blue outline-none"
                  />
                </div>
              </div>

              {/* Tabla Formal de Reporte por Empleado */}
              <div className="overflow-x-auto border border-slate-200 rounded-xl">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-100/90 text-slate-700 uppercase font-bold border-b border-slate-200 tracking-wider">
                    <tr>
                      <th className="py-3 px-3">Fecha</th>
                      <th className="py-3 px-3">Locker</th>
                      <th className="py-3 px-4">Colaborador</th>
                      <th className="py-3 px-3">Área</th>
                      <th className="py-3 px-3 bg-emerald-50/80 text-emerald-800 text-center font-bold border-l border-r border-slate-200">
                        1. Entrada
                      </th>
                      <th className="py-3 px-3 bg-amber-50/80 text-amber-800 text-center font-bold border-r border-slate-200">
                        2. Sal. Descanso
                      </th>
                      <th className="py-3 px-3 bg-blue-50/80 text-mango-blue text-center font-bold border-r border-slate-200">
                        3. Ret. Descanso
                      </th>
                      <th className="py-3 px-3 bg-rose-50/80 text-mango-red text-center font-bold">
                        4. Salida
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200 bg-white">
                    {matrizReporte.map((r) => (
                      <tr key={r.key} className="hover:bg-blue-50/30 transition">
                        <td className="py-2.5 px-3 font-mono text-slate-600 tabular-nums">
                          {r.fecha}
                        </td>
                        <td className="py-2.5 px-3 font-mono font-bold text-mango-blue tabular-nums">
                          {r.locker && r.locker !== 'S/N' ? (
                            `#${r.locker}`
                          ) : (
                            <span className="text-slate-400 font-normal">S/N</span>
                          )}
                        </td>
                        <td className="py-2.5 px-4 font-bold text-slate-900">{r.nombre}</td>
                        <td className="py-2.5 px-3 text-slate-600 font-medium">{r.area}</td>
                        <td
                          className={`py-2.5 px-3 text-center font-mono font-bold border-l border-r border-slate-100 tabular-nums ${
                            r.entrada !== '--:--'
                              ? 'text-emerald-700 bg-emerald-50/40'
                              : 'text-slate-300'
                          }`}
                        >
                          {r.entrada}
                        </td>
                        <td
                          className={`py-2.5 px-3 text-center font-mono font-bold border-r border-slate-100 tabular-nums ${
                            r.salidaDescanso !== '--:--'
                              ? 'text-amber-700 bg-amber-50/40'
                              : 'text-slate-300'
                          }`}
                        >
                          {r.salidaDescanso}
                        </td>
                        <td
                          className={`py-2.5 px-3 text-center font-mono font-bold border-r border-slate-100 tabular-nums ${
                            r.retornoDescanso !== '--:--'
                              ? 'text-mango-blue bg-blue-50/40'
                              : 'text-slate-300'
                          }`}
                        >
                          {r.retornoDescanso}
                        </td>
                        <td
                          className={`py-2.5 px-3 text-center font-mono font-bold tabular-nums ${
                            r.salida !== '--:--'
                              ? 'text-mango-red bg-rose-50/40'
                              : 'text-slate-300'
                          }`}
                        >
                          {r.salida}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {matrizReporte.length === 0 && (
                  <div className="py-10 text-center text-slate-400 text-xs">
                    <Inbox className="w-7 h-7 mx-auto mb-1.5 text-slate-300" />
                    No se encontraron registros de jornadas para los filtros seleccionados.
                  </div>
                )}
              </div>

              <div className="flex flex-col sm:flex-row items-center justify-between text-xs text-slate-500 pt-2 gap-1 font-medium">
                <span className="tabular-nums">
                  Total de jornadas en el reporte: {matrizReporte.length}
                </span>
                <span className="text-mango-blue font-bold">
                  Mango Center Hiperferretería &bull; Almacenamiento Local Seguro
                </span>
              </div>
            </div>
          </section>
        )}

        {/* VISTA 3: CONFIGURACIÓN Y PERSONAL (EXCLUSIVO SUPERVISOR) */}
        {activeTab === 'config' && perfilActual === 'supervisor' && (
          <section className="space-y-4 sm:space-y-6">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200/90 p-4 sm:p-6 space-y-4 sm:space-y-6">
              <div className="border-b border-slate-100 pb-4 flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h2 className="text-sm sm:text-lg font-bold text-mango-blueDark flex items-center gap-2">
                    <UploadCloud className="w-5 h-5 text-mango-red" /> Base de Datos del Personal Mango Center
                  </h2>
                  <p className="text-[11px] sm:text-xs text-slate-500 mt-0.5">
                    Gestione el personal individualmente o cargue nóminas en formato Excel (.xlsx).
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={abrirModalAgregarPersonal}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold text-white bg-mango-blue hover:bg-mango-blueDark rounded-xl shadow-sm transition active:scale-[0.98] cursor-pointer"
                  >
                    <UserPlus className="w-4 h-4" /> + Agregar Personal
                  </button>
                  <button
                    type="button"
                    onClick={descargarPlantillaExcel}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold text-mango-blue hover:text-mango-blueDark bg-blue-50 hover:bg-blue-100 rounded-xl border border-blue-200 transition cursor-pointer"
                  >
                    <Download className="w-4 h-4" /> Plantilla (.xlsx)
                  </button>
                </div>
              </div>

              {/* Zona Drag & Drop */}
              <div className="border-2 border-dashed border-slate-300 rounded-2xl p-6 sm:p-8 text-center bg-slate-50 hover:bg-blue-50/40 transition cursor-pointer relative">
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={handleExcelUpload}
                  className="absolute inset-0 opacity-0 cursor-pointer"
                />
                <div className="w-12 h-12 sm:w-14 sm:h-14 mx-auto mb-2 sm:mb-3 bg-white rounded-full flex items-center justify-center border border-slate-200 shadow-sm text-mango-blue">
                  <FileUp className="w-6 h-6 sm:w-7 sm:h-7" />
                </div>
                <p className="text-xs sm:text-sm font-bold text-slate-800">
                  Haga clic o arrastre aquí su archivo Excel (.xlsx)
                </p>
                <p className="text-[11px] sm:text-xs text-slate-500 mt-1">
                  Columnas requeridas: <strong>Nombre y Apellido</strong>,{' '}
                  <strong>Área de Trabajo</strong> y opcionalmente{' '}
                  <strong>Número de Locker</strong>
                </p>
              </div>

              {/* Lista de Empleados en Memoria Local */}
              <div className="space-y-3">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-2">
                    <Users className="w-4 h-4 text-mango-blue" /> Personal Registrado en Memoria Local ({personal.length})
                  </h3>
                  {personal.length > 0 && (
                    <button
                      type="button"
                      onClick={() =>
                        setConfirmDialog({
                          open: true,
                          title: 'Vaciar Nómina de Personal',
                          message:
                            '¿Está seguro de que desea eliminar toda la lista de colaboradores guardada localmente?',
                          confirmLabel: 'Sí, vaciar lista',
                          onConfirm: () => {
                            setPersonal([]);
                            setColaboradorActual(null);
                            setConfirmDialog(null);
                            showToast('Nómina de colaboradores vaciada', 'warning');
                          }
                        })
                      }
                      className="text-xs text-mango-red hover:text-mango-redDark font-bold transition cursor-pointer"
                    >
                      Vaciar Lista
                    </button>
                  )}
                </div>

                <div className="overflow-x-auto border border-slate-200 rounded-xl max-h-96 overflow-y-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-100 text-slate-700 uppercase font-bold sticky top-0 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5 px-4">Locker</th>
                        <th className="py-2.5 px-4">Nombre y Apellido</th>
                        <th className="py-2.5 px-4">Área de Trabajo</th>
                        <th className="py-2.5 px-4 text-center">Turno Base</th>
                        <th className="py-2.5 px-4 text-right">Acciones</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 bg-white">
                      {personal.map((p) => (
                        <tr key={p.id} className="hover:bg-slate-50 transition">
                          <td className="py-2.5 px-4 font-mono font-bold text-mango-blue tabular-nums">
                            {p.locker && p.locker !== 'S/N' ? (
                              `#${p.locker}`
                            ) : (
                              <span className="text-slate-400 font-semibold">
                                S/N (Sin Asignar)
                              </span>
                            )}
                          </td>
                          <td className="py-2.5 px-4 font-bold text-slate-800">{p.nombre}</td>
                          <td className="py-2.5 px-4 text-slate-600 font-medium">{p.area}</td>
                          <td className="py-2.5 px-4 text-center text-slate-600 font-mono">
                            {p.turnoBase && p.turnoBase !== 'auto'
                              ? `${p.turnoBase} AM`
                              : 'Auto-detectar'}
                          </td>
                          <td className="py-2.5 px-4 text-right">
                            <div className="inline-flex items-center gap-1">
                              <button
                                type="button"
                                onClick={() => abrirModalEditarPersonal(p)}
                                className="p-1.5 text-slate-500 hover:text-mango-blue rounded-lg hover:bg-blue-50 transition cursor-pointer"
                                title="Editar colaborador"
                              >
                                <Edit3 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => eliminarColaborador(p)}
                                className="p-1.5 text-slate-400 hover:text-mango-red rounded-lg hover:bg-rose-50 transition cursor-pointer"
                                title="Eliminar colaborador"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {personal.length === 0 && (
                    <div className="py-8 text-center text-slate-400 text-xs">
                      Aún no se ha cargado ninguna nómina de colaboradores.
                    </div>
                  )}
                </div>
              </div>
            </div>
          </section>
        )}

        {/* VISTA 4: MÓDULO DE GRÁFICOS DE RETARDO (> 5 MIN) (EXCLUSIVO SUPERVISOR) */}
        {activeTab === 'graficos' && perfilActual === 'supervisor' && (
          <section className="space-y-4 sm:space-y-6">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200/90 p-4 sm:p-6 space-y-4 sm:space-y-6">
              {/* Barra Superior de Gráficos */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
                <div className="flex items-center gap-3">
                  <div className="w-9 h-9 sm:w-11 sm:h-11 rounded-xl bg-rose-50 flex items-center justify-center text-mango-red border border-rose-200 shrink-0">
                    <ClockAlert className="w-5 h-5 sm:w-6 sm:h-6" />
                  </div>
                  <div>
                    <h2 className="text-sm sm:text-lg font-bold text-slate-900 leading-tight">
                      Módulo Analítico de Retardos del Personal
                    </h2>
                    <p className="text-[11px] sm:text-xs text-slate-500">
                      Condición: Personal con retardo superior a <strong>5 min</strong> respecto a
                      su hora de entrada (<strong>07:00am, 08:00am, 09:00am</strong>).
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => showToast('Gráficos de retardos sincronizados.', 'info')}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-lg border border-slate-300 transition active:scale-[0.98] cursor-pointer"
                  >
                    <RefreshCw className="w-3.5 h-3.5" /> Actualizar Gráficos
                  </button>
                </div>
              </div>

              {/* Filtros de Análisis de Retardos */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-slate-50 p-3 sm:p-4 rounded-xl border border-slate-200 text-xs">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Período de Análisis:
                  </label>
                  <select
                    value={filtroGraficoPeriodo}
                    onChange={(e) =>
                      setFiltroGraficoPeriodo(
                        e.target.value as 'diario' | 'semanal' | 'quincenal' | 'mensual' | 'todos'
                      )
                    }
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 font-medium focus:ring-2 focus:ring-mango-blue/20 focus:border-mango-blue outline-none"
                  >
                    <option value="diario">Diario (Hoy)</option>
                    <option value="semanal">Semanal (Últimos 7 días)</option>
                    <option value="quincenal">Quincenal (Últimos 15 días)</option>
                    <option value="mensual">Mensual (Últimos 30 días)</option>
                    <option value="todos">Histórico Completo</option>
                  </select>
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Turno de Entrada Evaluado:
                  </label>
                  <select
                    value={filtroGraficoTurno}
                    onChange={(e) =>
                      setFiltroGraficoTurno(
                        e.target.value as 'todos' | '07:00' | '08:00' | '09:00'
                      )
                    }
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 font-medium focus:ring-2 focus:ring-mango-blue/20 focus:border-mango-blue outline-none"
                  >
                    <option value="todos">Todos los Turnos (07:00, 08:00, 09:00)</option>
                    <option value="07:00">Turno 07:00 AM (Tolerancia máx: 07:05 AM)</option>
                    <option value="08:00">Turno 08:00 AM (Tolerancia máx: 08:05 AM)</option>
                    <option value="09:00">Turno 09:00 AM (Tolerancia máx: 09:05 AM)</option>
                  </select>
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Filtrar por Área / Departamento:
                  </label>
                  <select
                    value={filtroGraficoArea}
                    onChange={(e) => setFiltroGraficoArea(e.target.value)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 font-medium focus:ring-2 focus:ring-mango-blue/20 focus:border-mango-blue outline-none"
                  >
                    <option value="">Todas las áreas</option>
                    {areasDisponibles.map((a) => (
                      <option key={a} value={a}>
                        {a}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Tarjetas Resumen (KPIs) de Retardos */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4">
                <div className="bg-blue-50/60 border border-blue-200/80 rounded-xl p-3 sm:p-4 text-center">
                  <span className="text-[10px] sm:text-xs font-bold text-mango-blue block uppercase">
                    Entradas Evaluadas
                  </span>
                  <span className="text-lg sm:text-2xl font-extrabold text-slate-900 mt-0.5 block tabular-nums">
                    {datosAnaliticaRetardos.totalEvaluadas}
                  </span>
                  <span className="text-[10px] text-slate-500">Total en el período</span>
                </div>

                <div className="bg-emerald-50/60 border border-emerald-200/80 rounded-xl p-3 sm:p-4 text-center">
                  <span className="text-[10px] sm:text-xs font-bold text-emerald-800 block uppercase">
                    Puntuales / Tolerancia
                  </span>
                  <span className="text-lg sm:text-2xl font-extrabold text-emerald-700 mt-0.5 block tabular-nums">
                    {datosAnaliticaRetardos.puntuales}
                  </span>
                  <span className="text-[10px] text-emerald-600 font-medium">
                    &le; 5 min de retardo
                  </span>
                </div>

                <div className="bg-rose-50/70 border border-rose-200 rounded-xl p-3 sm:p-4 text-center">
                  <span className="text-[10px] sm:text-xs font-bold text-mango-red block uppercase">
                    Retardo Crítico
                  </span>
                  <span className="text-lg sm:text-2xl font-extrabold text-mango-red mt-0.5 block tabular-nums">
                    {datosAnaliticaRetardos.casosRetardoCritico.length}
                  </span>
                  <span className="text-[10px] text-rose-600 font-bold">
                    &gt; 5 min de demora
                  </span>
                </div>

                <div className="bg-amber-50/60 border border-amber-200/80 rounded-xl p-3 sm:p-4 text-center">
                  <span className="text-[10px] sm:text-xs font-bold text-amber-800 block uppercase">
                    Promedio Retardo
                  </span>
                  <span className="text-lg sm:text-2xl font-extrabold text-amber-700 mt-0.5 block tabular-nums">
                    {datosAnaliticaRetardos.promMinutos} min
                  </span>
                  <span className="text-[10px] text-amber-600">En casos &gt; 5 min</span>
                </div>
              </div>

              {/* Contenedores de Gráficos */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 sm:gap-6">
                {/* Gráfico 1: Barras de Colaboradores con Retardo > 5 min */}
                <div className="lg:col-span-8 bg-slate-50/50 border border-slate-200 rounded-xl p-4 sm:p-5">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-1.5">
                      <BarChart2 className="w-4 h-4 text-mango-red" /> Minutos de Demora Acumulados
                      por Colaborador (&gt; 5 min)
                    </h3>
                    <span className="text-[11px] font-bold text-mango-red">
                      Exclusivo &gt; 5 min
                    </span>
                  </div>

                  {datosAnaliticaRetardos.barrasColaboradores.length === 0 ? (
                    <div className="text-center text-xs text-slate-400 py-16 space-y-2">
                      <p>
                        ¡Excelente! No hay colaboradores con retardo mayor a 5 minutos en el
                        período seleccionado.
                      </p>
                      <button
                        type="button"
                        onClick={generarDatosDemostrativos}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-mango-blue bg-blue-50 hover:bg-blue-100 border border-blue-200 transition cursor-pointer"
                      >
                        <Sparkles className="w-3.5 h-3.5" /> Cargar casos de ejemplo para visualizar
                      </button>
                    </div>
                  ) : (
                    <div className="space-y-3 py-2">
                      {(() => {
                        const maxMin = Math.max(
                          ...datosAnaliticaRetardos.barrasColaboradores.map((b) => b.minutos),
                          30
                        );
                        return datosAnaliticaRetardos.barrasColaboradores.map((item) => {
                          const widthPct = Math.max(Math.round((item.minutos / maxMin) * 100), 8);
                          return (
                            <div key={item.label} className="space-y-1">
                              <div className="flex items-center justify-between text-xs">
                                <span className="font-bold text-slate-800">{item.label}</span>
                                <span className="font-mono font-bold text-mango-red tabular-nums">
                                  +{item.minutos} min
                                </span>
                              </div>
                              <div className="w-full h-7 bg-slate-200/80 rounded-lg overflow-hidden p-0.5">
                                <div
                                  style={{ width: `${widthPct}%` }}
                                  className="h-full bg-mango-red rounded-md transition-all duration-300 flex items-center justify-end pr-2"
                                >
                                  <span className="text-[10px] font-mono font-bold text-white tabular-nums">
                                    {item.minutos}m
                                  </span>
                                </div>
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                  )}
                </div>

                {/* Gráfico 2: Distribución por Turno de Entrada (07:00, 08:00, 09:00) */}
                <div className="lg:col-span-4 bg-slate-50/50 border border-slate-200 rounded-xl p-4 sm:p-5 flex flex-col justify-between">
                  <div className="mb-3">
                    <h3 className="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-1.5">
                      <PieChart className="w-4 h-4 text-mango-blue" /> Retardos por Turno de Entrada
                    </h3>
                    <p className="text-[10px] text-slate-500">07:00 AM | 08:00 AM | 09:00 AM</p>
                  </div>

                  <div className="my-auto py-4 flex flex-col items-center">
                    {totalTurnosRetardo === 0 ? (
                      <div className="text-center text-xs text-slate-400 py-8">
                        Sin casos de retardo por turno
                      </div>
                    ) : (
                      <>
                        {/* Gráfico Donut SVG */}
                        <div className="relative w-40 h-40 flex items-center justify-center">
                          <svg viewBox="0 0 36 36" className="w-full h-full -rotate-90">
                            <circle
                              cx="18"
                              cy="18"
                              r="15.915"
                              fill="transparent"
                              stroke="#e2e8f0"
                              strokeWidth="4.2"
                            />
                            {(() => {
                              const segments = [
                                {
                                  label: '07:00 AM',
                                  val: datosAnaliticaRetardos.conteoTurnos['07:00 AM'],
                                  color: '#e5242d'
                                },
                                {
                                  label: '08:00 AM',
                                  val: datosAnaliticaRetardos.conteoTurnos['08:00 AM'],
                                  color: '#0d5ca4'
                                },
                                {
                                  label: '09:00 AM',
                                  val: datosAnaliticaRetardos.conteoTurnos['09:00 AM'],
                                  color: '#d97706'
                                }
                              ];
                              let accumulatedPct = 0;
                              return segments.map((seg) => {
                                if (seg.val === 0) return null;
                                const pct = (seg.val / totalTurnosRetardo) * 100;
                                const strokeDasharray = `${pct} ${100 - pct}`;
                                const strokeDashoffset = -accumulatedPct;
                                accumulatedPct += pct;
                                return (
                                  <circle
                                    key={seg.label}
                                    cx="18"
                                    cy="18"
                                    r="15.915"
                                    fill="transparent"
                                    stroke={seg.color}
                                    strokeWidth="4.2"
                                    strokeDasharray={strokeDasharray}
                                    strokeDashoffset={strokeDashoffset}
                                  />
                                );
                              });
                            })()}
                          </svg>
                          <div className="absolute inset-0 flex flex-col items-center justify-center text-center">
                            <span className="text-xl font-extrabold text-slate-900 tabular-nums">
                              {totalTurnosRetardo}
                            </span>
                            <span className="text-[10px] font-bold text-slate-500 uppercase">
                              Casos
                            </span>
                          </div>
                        </div>
                      </>
                    )}
                  </div>

                  {/* Leyenda de turnos */}
                  <div className="space-y-1.5 pt-3 border-t border-slate-200/80 text-xs">
                    <div className="flex items-center justify-between">
                      <span className="flex items-center gap-2 font-medium text-slate-700">
                        <span className="w-3 h-3 rounded-xs bg-mango-red inline-block" />
                        Turno 07:00 AM
                      </span>
                      <span className="font-mono font-bold text-slate-900 tabular-nums">
                        {datosAnaliticaRetardos.conteoTurnos['07:00 AM']}
                      </span>
                    </div>
                    <div className="flex items-center justify-between">
                      <span className="flex items-center gap-2 font-medium text-slate-700">
                        <span className="w-3 h-3 rounded-xs bg-mango-blue inline-block" />
                        Turno 08:00 AM
                      </span>
                      <span className="font-mono font-bold text-slate-900 tabular-nums">
                        {datosAnaliticaRetardos.conteoTurnos['08:00 AM']}
                      </span>
                    </div>
                    <div className="flex items-center justify-between">
                      <span className="flex items-center gap-2 font-medium text-slate-700">
                        <span className="w-3 h-3 rounded-xs bg-amber-600 inline-block" />
                        Turno 09:00 AM
                      </span>
                      <span className="font-mono font-bold text-slate-900 tabular-nums">
                        {datosAnaliticaRetardos.conteoTurnos['09:00 AM']}
                      </span>
                    </div>
                  </div>
                </div>
              </div>

              {/* Tabla Detallada Formal de Personal con Retardo > 5 min */}
              <div className="space-y-2 pt-2">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-2">
                    <ListFilter className="w-4 h-4 text-mango-red" /> Detalle Nominal de Casos con
                    Retardo (&gt; 5 min)
                  </h3>
                  <span className="text-xs font-bold text-mango-red tabular-nums">
                    {datosAnaliticaRetardos.casosRetardoCritico.length} caso(s) &gt; 5 min
                  </span>
                </div>

                <div className="overflow-x-auto border border-slate-200 rounded-xl">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead className="bg-slate-100 text-slate-700 uppercase font-bold border-b border-slate-200 tracking-wider">
                      <tr>
                        <th className="py-2.5 px-3">Fecha</th>
                        <th className="py-2.5 px-3">Locker</th>
                        <th className="py-2.5 px-3">Colaborador</th>
                        <th className="py-2.5 px-3">Área</th>
                        <th className="py-2.5 px-3 text-center">Turno Base</th>
                        <th className="py-2.5 px-3 text-center">Hora Llegada</th>
                        <th className="py-2.5 px-3 text-center bg-rose-50 text-mango-red font-extrabold">
                          Minutos Retardo
                        </th>
                        <th className="py-2.5 px-3 text-center">Condición</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200 bg-white">
                      {datosAnaliticaRetardos.casosRetardoCritico.map((r) => (
                        <tr
                          key={r.id}
                          className="hover:bg-rose-50/40 transition border-b border-slate-100"
                        >
                          <td className="py-2.5 px-3 font-mono text-slate-600 tabular-nums">
                            {r.fecha}
                          </td>
                          <td className="py-2.5 px-3 font-mono font-bold text-mango-blue tabular-nums">
                            {r.locker && r.locker !== 'S/N' ? `#${r.locker}` : 'S/N'}
                          </td>
                          <td className="py-2.5 px-3 font-bold text-slate-900">{r.nombre}</td>
                          <td className="py-2.5 px-3 text-slate-600">{r.area}</td>
                          <td className="py-2.5 px-3 text-center font-bold text-slate-700 tabular-nums">
                            {r.turno}
                          </td>
                          <td className="py-2.5 px-3 text-center font-mono font-bold text-rose-700 tabular-nums">
                            {r.hora}
                          </td>
                          <td className="py-2.5 px-3 text-center font-mono font-extrabold text-mango-red bg-rose-50 tabular-nums">
                            +{r.minutosRetardo} min
                          </td>
                          <td className="py-2.5 px-3 text-center font-bold text-rose-800">
                            Retardo &gt; 5 min
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {datosAnaliticaRetardos.casosRetardoCritico.length === 0 && (
                    <div className="py-8 text-center text-slate-400 text-xs">
                      <CheckCircle className="w-7 h-7 mx-auto mb-1.5 text-emerald-500" />
                      No se registraron retardos superiores a 5 minutos para los filtros
                      seleccionados.
                    </div>
                  )}
                </div>
              </div>
            </div>
          </section>
        )}
      </main>

      {/* MODAL: AGREGAR / EDITAR PERSONAL */}
      {modalNuevoPersonalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 no-print">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-md w-full p-5 sm:p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-mango-blue/10 flex items-center justify-center text-mango-blue border border-mango-blue/20">
                  <UserPlus className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm sm:text-base font-bold text-slate-900">
                    {editingColab ? 'Editar Personal' : 'Agregar Nuevo Personal'}
                  </h3>
                  <p className="text-[11px] sm:text-xs text-slate-500">
                    Mango Center Hiperferretería
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalNuevoPersonalOpen(false)}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-lg hover:bg-slate-100 transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={guardarPersonal} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">
                  Nombre y Apellidos <span className="text-mango-red">*</span>
                </label>
                <input
                  type="text"
                  value={nuevoNombre}
                  onChange={(e) => setNuevoNombre(e.target.value)}
                  placeholder="Ej. José Ramón Pérez..."
                  required
                  autoFocus
                  className="w-full text-xs sm:text-sm px-3.5 py-2.5 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:border-mango-blue focus:ring-2 focus:ring-mango-blue/20 outline-none"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">
                  Área o Departamento <span className="text-mango-red">*</span>
                </label>
                <input
                  type="text"
                  value={nuevaArea}
                  onChange={(e) => setNuevaArea(e.target.value)}
                  placeholder="Ej. Piso de Venta, Almacén, Cajas..."
                  required
                  list="listaSugerenciasAreas"
                  className="w-full text-xs sm:text-sm px-3.5 py-2.5 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:border-mango-blue focus:ring-2 focus:ring-mango-blue/20 outline-none"
                />
                <datalist id="listaSugerenciasAreas">
                  <option value="Piso de Venta / Ferretería" />
                  <option value="Cajas y Facturación" />
                  <option value="Almacén Central / Despacho" />
                  <option value="Atención al Cliente" />
                  <option value="Administración" />
                  <option value="Seguridad / Control" />
                  <option value="Mantenimiento" />
                </datalist>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="block text-xs font-bold text-slate-700">
                      N° de Locker
                    </label>
                    <span className="text-[10px] font-semibold text-emerald-700">Opcional</span>
                  </div>
                  <input
                    type="text"
                    value={nuevoLocker}
                    onChange={(e) => setNuevoLocker(e.target.value)}
                    placeholder="Ej. 105 o vacío"
                    className="w-full text-xs sm:text-sm px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:border-mango-blue outline-none uppercase font-mono"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Turno de Entrada
                  </label>
                  <select
                    value={nuevoTurnoBase}
                    onChange={(e) =>
                      setNuevoTurnoBase(
                        e.target.value as 'auto' | '07:00' | '08:00' | '09:00'
                      )
                    }
                    className="w-full text-xs sm:text-sm px-3 py-2 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:border-mango-blue outline-none"
                  >
                    <option value="auto">Automático</option>
                    <option value="07:00">07:00 AM</option>
                    <option value="08:00">08:00 AM</option>
                    <option value="09:00">09:00 AM</option>
                  </select>
                </div>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setModalNuevoPersonalOpen(false)}
                  className="px-3.5 py-2 text-xs font-bold text-slate-600 hover:text-slate-800 bg-slate-100 hover:bg-slate-200 rounded-xl transition cursor-pointer"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 text-xs font-bold text-white bg-mango-blue hover:bg-mango-blueDark rounded-xl shadow-md shadow-mango-blue/20 transition active:scale-[0.98] cursor-pointer"
                >
                  Guardar
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: LOGIN SUPERVISOR */}
      {modalLoginSupervisorOpen && (
        <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 no-print">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-sm w-full p-5 sm:p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 sm:w-10 sm:h-10 rounded-xl bg-amber-500/10 flex items-center justify-center text-amber-600 border border-amber-500/20">
                  <Lock className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm sm:text-base font-bold text-slate-900">
                    Acceso Supervisor
                  </h3>
                  <p className="text-[11px] text-slate-500">
                    Módulos Administrativos y Gráficos
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalLoginSupervisorOpen(false)}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-lg hover:bg-slate-100 transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={verificarClaveSupervisor} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">
                  Clave de Seguridad
                </label>
                <input
                  type="password"
                  value={claveSupervisorInput}
                  onChange={(e) => setClaveSupervisorInput(e.target.value)}
                  placeholder="Ingrese su clave..."
                  required
                  autoFocus
                  className="w-full text-xs sm:text-sm px-3.5 py-2.5 bg-slate-50 border border-slate-300 rounded-xl focus:bg-white focus:border-mango-blue outline-none"
                />
                <p className="text-[11px] text-slate-400 mt-1">
                  Clave inicial por defecto: <strong className="text-slate-600">1234</strong>
                </p>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setModalLoginSupervisorOpen(false)}
                  className="px-3.5 py-2 text-xs font-bold text-slate-600 bg-slate-100 rounded-xl cursor-pointer"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 text-xs font-bold text-white bg-mango-blue hover:bg-mango-blueDark rounded-xl shadow-md transition cursor-pointer"
                >
                  Desbloquear
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: CAMBIAR CLAVE SUPERVISOR */}
      {modalCambiarClaveOpen && (
        <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 no-print">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-sm w-full p-5 sm:p-6 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 rounded-xl bg-blue-50 flex items-center justify-center text-mango-blue border border-blue-200">
                  <KeyRound className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm sm:text-base font-bold text-slate-900">Cambiar Clave</h3>
                  <p className="text-[11px] text-slate-500">Contraseña de supervisor</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalCambiarClaveOpen(false)}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-lg hover:bg-slate-100 transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={guardarNuevaClaveSupervisor} className="space-y-3">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">
                  Clave Actual
                </label>
                <input
                  type="password"
                  value={claveActualInput}
                  onChange={(e) => setClaveActualInput(e.target.value)}
                  required
                  autoFocus
                  className="w-full text-xs px-3 py-2.5 bg-slate-50 border border-slate-300 rounded-xl outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">
                  Nueva Clave (mín. 4 caracteres)
                </label>
                <input
                  type="password"
                  value={nuevaClaveInput}
                  onChange={(e) => setNuevaClaveInput(e.target.value)}
                  required
                  minLength={4}
                  className="w-full text-xs px-3 py-2.5 bg-slate-50 border border-slate-300 rounded-xl outline-none"
                />
              </div>
              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setModalCambiarClaveOpen(false)}
                  className="px-3.5 py-2 text-xs font-bold text-slate-600 bg-slate-100 rounded-xl cursor-pointer"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 text-xs font-bold text-white bg-mango-blue rounded-xl cursor-pointer"
                >
                  Actualizar
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL DE CONFIRMACIÓN PERSONALIZADO */}
      {confirmDialog?.open && (
        <div className="fixed inset-0 bg-slate-900/70 backdrop-blur-xs z-50 flex items-center justify-center p-3 sm:p-4 no-print">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-sm w-full p-5 sm:p-6 space-y-4">
            <h3 className="text-base font-bold text-slate-900">{confirmDialog.title}</h3>
            <p className="text-xs text-slate-600 leading-relaxed">{confirmDialog.message}</p>
            <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setConfirmDialog(null)}
                className="px-3.5 py-2 text-xs font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 rounded-xl cursor-pointer"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={confirmDialog.onConfirm}
                className="px-4 py-2 text-xs font-bold text-white bg-mango-red hover:bg-mango-redDark rounded-xl shadow-sm cursor-pointer"
              >
                {confirmDialog.confirmLabel}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ÁREA DE IMPRESIÓN FORMAL MEMBRETADA MANGO CENTER */}
      <div id="printArea" className="p-8 hidden">
        <div className="border-b-4 border-mango-red pb-4 mb-6 flex justify-between items-center bg-slate-900 text-white p-4 rounded-lg">
          <div className="flex items-center gap-4">
            <img
              src="/assets/aistudio/mango_logo.jpg"
              alt="Mango Center"
              className="h-12 w-auto object-contain"
            />
            <div>
              <h1 className="text-lg font-extrabold uppercase tracking-wider text-white">
                MANGO CENTER HIPERFERRETERÍA
              </h1>
              <p className="text-xs text-slate-300">
                Reporte Consolidado de Asistencia y Control de Jornadas
              </p>
              <p className="text-[11px] text-sky-400 font-semibold mt-0.5">
                Período: {filtroPeriodo.toUpperCase()} | Total registrados: {matrizReporte.length}
              </p>
            </div>
          </div>
          <div className="text-right text-xs text-slate-300">
            <p>
              Fecha de emisión:{' '}
              <strong className="text-white">{now.toLocaleString('es-ES')}</strong>
            </p>
            <p>Departamento de Gestión Humana</p>
          </div>
        </div>

        <table className="w-full text-left text-xs border border-slate-300">
          <thead className="bg-slate-100 text-slate-800 uppercase font-bold border-b border-slate-300">
            <tr>
              <th className="py-2 px-2 border-r border-slate-300">Fecha</th>
              <th className="py-2 px-2 border-r border-slate-300">Locker</th>
              <th className="py-2 px-3 border-r border-slate-300">Colaborador</th>
              <th className="py-2 px-2 border-r border-slate-300">Área</th>
              <th className="py-2 px-2 border-r border-slate-300 text-center">1. Entrada</th>
              <th className="py-2 px-2 border-r border-slate-300 text-center">
                2. Sal. Descanso
              </th>
              <th className="py-2 px-2 border-r border-slate-300 text-center">
                3. Ret. Descanso
              </th>
              <th className="py-2 px-2 text-center">4. Salida</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200">
            {matrizReporte.map((r) => (
              <tr key={r.key}>
                <td className="py-1.5 px-2 border-r border-slate-200 font-mono">{r.fecha}</td>
                <td className="py-1.5 px-2 border-r border-slate-200 font-bold text-slate-800">
                  {r.locker && r.locker !== 'S/N' ? `#${r.locker}` : 'S/N'}
                </td>
                <td className="py-1.5 px-3 border-r border-slate-200 font-semibold">
                  {r.nombre}
                </td>
                <td className="py-1.5 px-2 border-r border-slate-200">{r.area}</td>
                <td className="py-1.5 px-2 border-r border-slate-200 text-center font-mono font-bold">
                  {r.entrada}
                </td>
                <td className="py-1.5 px-2 border-r border-slate-200 text-center font-mono font-bold">
                  {r.salidaDescanso}
                </td>
                <td className="py-1.5 px-2 border-r border-slate-200 text-center font-mono font-bold">
                  {r.retornoDescanso}
                </td>
                <td className="py-1.5 px-2 text-center font-mono font-bold">{r.salida}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <div className="mt-14 grid grid-cols-2 gap-12 text-center text-xs text-slate-700">
          <div className="border-t-2 border-slate-400 pt-3">
            <p className="font-bold text-slate-900">Supervisor de Turno / Seguridad</p>
            <p className="text-slate-500">Mango Center Hiperferretería</p>
          </div>
          <div className="border-t-2 border-slate-400 pt-3">
            <p className="font-bold text-slate-900">Gerencia de Talento Humano</p>
            <p className="text-slate-500">Firma y Sello de Conformidad</p>
          </div>
        </div>
      </div>

      {/* CINTILLO INFERIOR SIEMPRE VISIBLE CON LOGO S&S WEB STUDIO */}
      <footer className="sticky bottom-0 z-40 bg-mango-dark border-t-2 border-mango-blue text-slate-300 text-[11px] sm:text-xs py-2 px-3 sm:px-6 no-print mt-auto shadow-lg">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-2">
          <span className="font-medium text-slate-300 flex items-center gap-1.5">
            <Wrench className="w-3.5 h-3.5 text-mango-red shrink-0" />
            <span>Mango Center Hiperferretería &copy; Control de Asistencia y Jornada</span>
          </span>
          <div className="flex items-center gap-2.5 bg-white/95 text-slate-900 px-3 py-1 rounded-xl shadow-sm border border-sky-400/40">
            <img
              src="/assets/aistudio/IMG-20261013-WA0042_ee5e03ef.jpg"
              alt="S&S Web Studio"
              className="h-6 sm:h-7 w-auto object-contain rounded-xs shrink-0"
            />
            <span className="text-[11px] sm:text-xs font-extrabold tracking-tight text-slate-900 whitespace-nowrap">
              Creado Diseñado por S&amp;S Web Studio - 0424-5104245
            </span>
          </div>
        </div>
      </footer>
    </div>
  );
}
