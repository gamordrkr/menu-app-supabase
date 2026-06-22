import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { createPortal } from 'react-dom';
import {
  RefreshCw,
  Download,
  Check,
  AlertTriangle,
  Shuffle,
  X,
  Edit3,
  Plus,
  Loader2,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  listClientes,
  createCliente,
  listPlatillos,
  bulkInsertPlatillos,
  getLastCiclo,
  saveCiclo,
} from './supabaseClient.js';

const CATEGORIAS = [
  'AVENA CALIENTE',
  'GUARNICION 1',
  'GUARNICION 2',
  'PLATO FUERTE HUEVO',
  'PLATO FUERTE MEXICANO O DULCE',
  'PLATO FUERTE VEGETARIANO / VEGANO',
  'COMPLEMENTOS',
];

const FIXED_CATEGORIES = {
  'AVENA CALIENTE': 'Avena sencilla',
  COMPLEMENTOS: 'Complementos',
};

const CATEGORIA_LABEL = {
  'AVENA CALIENTE': 'AVENA',
  'GUARNICION 1': 'GUARNICION 1',
  'GUARNICION 2': 'GUARNICION 2',
  'PLATO FUERTE HUEVO': 'PLATO FUERTE HUEVO',
  'PLATO FUERTE MEXICANO O DULCE': 'PLATO FUERTE MEXICANO O DULCE',
  'PLATO FUERTE VEGETARIANO / VEGANO': 'PLATO FUERTE VEGETARIANO / VEGANO',
  COMPLEMENTOS: 'COMPLEMENTOS (PAN, TORTILLA, SALSA Y LIMONES)',
};

// Versión corta solo para la columna de categoría de la tabla en pantalla.
// El Excel exportado y el modal de edición siguen usando CATEGORIA_LABEL completo.
const CATEGORIA_LABEL_TABLA = {
  'AVENA CALIENTE': 'AVENA',
  'GUARNICION 1': 'GUARNICION 1',
  'GUARNICION 2': 'GUARNICION 2',
  'PLATO FUERTE HUEVO': 'P. FUERTE HUEVO',
  'PLATO FUERTE MEXICANO O DULCE': 'P. FUERTE MEX/DULCE',
  'PLATO FUERTE VEGETARIANO / VEGANO': 'P. FUERTE VEGETARIANO',
  COMPLEMENTOS: 'COMPLEMENTOS',
};

const DIAS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'];

// Un único estilo de advertencia para toda celda en conflicto de cooldown.
// (Antes había una paleta de varios colores por familia; resultaba confusa
// sin dar información real, así que se simplificó a un solo color + modal con detalle.)
const CONFLICT_STYLE = { bg: '#FAEEDA', border: '#BA7517', text: '#633806' };

const COOLDOWN_WEEKS = 4; // misma familia no puede repetirse antes de N semanas

function buildIndexByCategory(catalog) {
  const map = {};
  for (const cat of CATEGORIAS) map[cat] = [];
  for (const d of catalog) {
    if (map[d.categoria]) map[d.categoria].push(d);
  }
  return map;
}

function mulberry32(a) {
  return function () {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function shuffle(arr, rng) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// Genera el ciclo respetando cooldown de COOLDOWN_WEEKS entre apariciones
// de la misma familia_id, dentro de la misma categoría, a lo largo de TODO
// el ciclo (no solo dentro de una semana).
function generateCycle(catalog, numSemanas, seed) {
  const byCat = buildIndexByCategory(catalog);
  const rng = mulberry32(seed);
  const cycle = [];

  // lastSeenWeek[categoria][familia_id] = índice de semana donde apareció por última vez
  const lastSeenWeek = {};
  for (const cat of CATEGORIAS) lastSeenWeek[cat] = {};

  for (let s = 0; s < numSemanas; s++) {
    const semana = DIAS.map(() => ({}));
    for (const cat of CATEGORIAS) {
      if (FIXED_CATEGORIES[cat]) {
        const fixedDish = {
          categoria: cat,
          platillo: FIXED_CATEGORIES[cat],
          tipo: '',
          ingredientes: '',
          alergeno: '',
          kcal: '',
          id: `fixed_${cat}`,
          keywords: [],
          familia_id: `fixed_${cat}`,
        };
        for (let dayIdx = 0; dayIdx < 5; dayIdx++) semana[dayIdx][cat] = fixedDish;
        continue;
      }

      const pool = shuffle([...byCat[cat]], rng);
      const usedNamesThisWeek = new Set();
      const usedFamiliesThisWeek = new Set();
      const chosen = [];

      for (let dayIdx = 0; dayIdx < 5; dayIdx++) {
        // Nivel 1: familia respeta cooldown global Y no se repitió esta semana
        let pick = pool.find((d) => {
          const lastWeek = lastSeenWeek[cat][d.familia_id];
          const cooledDown = lastWeek === undefined || s - lastWeek >= COOLDOWN_WEEKS;
          return (
            cooledDown &&
            !usedNamesThisWeek.has(d.platillo) &&
            !usedFamiliesThisWeek.has(d.familia_id)
          );
        });
        // Nivel 2: relaja cooldown pero mantiene no-repetir-en-semana
        if (!pick) {
          pick = pool.find(
            (d) =>
              !usedNamesThisWeek.has(d.platillo) &&
              !usedFamiliesThisWeek.has(d.familia_id)
          );
        }
        // Nivel 3: solo evita nombre exacto
        if (!pick) pick = pool.find((d) => !usedNamesThisWeek.has(d.platillo));
        // Último recurso
        if (!pick) pick = pool[dayIdx % pool.length];

        chosen.push(pick);
        usedNamesThisWeek.add(pick.platillo);
        usedFamiliesThisWeek.add(pick.familia_id);
        const idx = pool.indexOf(pick);
        if (idx >= 0) pool.splice(idx, 1);
      }

      chosen.forEach((dish, dayIdx) => {
        semana[dayIdx][cat] = dish;
        lastSeenWeek[cat][dish.familia_id] = s;
      });
    }
    cycle.push(semana);
  }
  return cycle;
}

// Para cada celda dentro de COOLDOWN_WEEKS de otra con la misma familia,
// devuelve { highlight: bool, conflicts: [{semanaIdx, diaIdx, platillo}] }
// para que el modal pueda mostrar info real: con qué otro platillo choca y dónde.
function computeViolationHighlights(cycle) {
  // occurrences[categoria][familia_id] = [{semanaIdx, diaIdx, platillo}, ...]
  const occurrences = {};
  for (const cat of CATEGORIAS) occurrences[cat] = {};

  for (let s = 0; s < cycle.length; s++) {
    for (const cat of CATEGORIAS) {
      if (FIXED_CATEGORIES[cat]) continue;
      for (let d = 0; d < 5; d++) {
        const dish = cycle[s][d][cat];
        if (!dish) continue;
        if (!occurrences[cat][dish.familia_id]) occurrences[cat][dish.familia_id] = [];
        occurrences[cat][dish.familia_id].push({ semanaIdx: s, diaIdx: d, platillo: dish.platillo });
      }
    }
  }

  // highlights[semanaIdx][diaIdx][categoria] = { conflicts: [...] } | null
  const highlights = cycle.map(() => DIAS.map(() => ({})));

  for (const cat of CATEGORIAS) {
    if (FIXED_CATEGORIES[cat]) {
      cycle.forEach((_, s) => DIAS.forEach((__, d) => (highlights[s][d][cat] = null)));
      continue;
    }
    for (const famId in occurrences[cat]) {
      const occs = occurrences[cat][famId];
      occs.forEach((occ, i) => {
        const conflicts = occs.filter(
          (other, j) => j !== i && Math.abs(other.semanaIdx - occ.semanaIdx) < COOLDOWN_WEEKS
        );
        highlights[occ.semanaIdx][occ.diaIdx][cat] =
          conflicts.length > 0 ? { conflicts } : null;
      });
    }
  }
  return highlights;
}

export default function MenuGenerator() {
  const [clients, setClients] = useState({}); // { [id]: { name, catalog } }
  const [activeClient, setActiveClient] = useState(null);
  const [numSemanas, setNumSemanas] = useState(4);
  const [seed, setSeed] = useState(1);
  const [cycle, setCycle] = useState(null);
  const [editing, setEditing] = useState(null);
  const [infoModal, setInfoModal] = useState(null); // { platillo, conflicts } | null
  const [loaded, setLoaded] = useState(false);
  const [showNewClient, setShowNewClient] = useState(false);
  const [saveStatus, setSaveStatus] = useState('idle'); // idle | saving | saved | error
  const [loadError, setLoadError] = useState('');
  const [exporting, setExporting] = useState(false);
  const [toolbarCollapsed, setToolbarCollapsed] = useState(false);
  const catalog = clients[activeClient]?.catalog || [];
  const byCat = useMemo(() => buildIndexByCategory(catalog), [catalog]);

  // --- Colapsa la barra superior al deslizar hacia abajo. Solo vuelve a
  // abrirse con click en el botón (no se auto-restaura al subir el scroll).
  useEffect(() => {
    let lastY = window.scrollY;
    let ticking = false;
    const THRESHOLD = 8;

    const onScroll = () => {
      if (ticking) return;
      ticking = true;
      requestAnimationFrame(() => {
        const currentY = window.scrollY;
        const delta = currentY - lastY;
        if (delta > THRESHOLD && currentY > 80) {
          setToolbarCollapsed(true);
          lastY = currentY;
        } else if (Math.abs(delta) > THRESHOLD) {
          lastY = currentY;
        }
        ticking = false;
      });
    };

    window.addEventListener('scroll', onScroll, { passive: true });
    return () => window.removeEventListener('scroll', onScroll);
  }, []);

  // --- Carga inicial: lista de clientes desde Supabase ---
  useEffect(() => {
    (async () => {
      try {
        const rows = await listClientes();
        if (rows.length === 0) {
          setLoaded(true);
          return;
        }
        const clientMap = {};
        for (const c of rows) {
          clientMap[c.id] = { name: c.nombre, catalog: null }; // catalog se carga al seleccionar
        }
        setClients(clientMap);
        const firstId = rows[0].id;
        await selectClient(firstId, clientMap);
      } catch (e) {
        setLoadError(
          'No se pudo conectar a la base de datos. Revisa la configuración de Supabase.'
        );
        setLoaded(true);
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const selectClient = async (clientId, clientMapOverride) => {
    setLoaded(false);
    try {
      const cat = await listPlatillos(clientId);
      const baseMap = clientMapOverride || clients;
      setClients({
        ...baseMap,
        [clientId]: { ...baseMap[clientId], catalog: cat },
      });
      setActiveClient(clientId);

      const lastCiclo = await getLastCiclo(clientId);
      if (lastCiclo) {
        setCycle(lastCiclo.datos);
        setNumSemanas(lastCiclo.num_semanas);
        setSeed(lastCiclo.seed);
      } else if (cat.length > 0) {
        setCycle(generateCycle(cat, numSemanas, seed));
      } else {
        setCycle([]);
      }
    } catch (e) {
      setLoadError('No se pudo cargar el catálogo de este cliente.');
    }
    setLoaded(true);
  };

  // --- Guarda el ciclo en Supabase cada vez que cambia ---
  useEffect(() => {
    if (!loaded || !cycle || cycle.length === 0 || !activeClient) return;
    setSaveStatus('saving');
    const t = setTimeout(async () => {
      try {
        await saveCiclo(activeClient, numSemanas, seed, cycle);
        setSaveStatus('saved');
      } catch (e) {
        setSaveStatus('error');
      }
    }, 500);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cycle]);

  const highlights = useMemo(
    () => (cycle ? computeViolationHighlights(cycle) : []),
    [cycle]
  );

  const handleClientChange = (clientId) => {
    selectClient(clientId);
  };

  const handleRegenerate = useCallback(() => {
    const newSeed = seed + 1;
    setSeed(newSeed);
    setCycle(generateCycle(catalog, numSemanas, newSeed));
  }, [seed, numSemanas, catalog]);

  const handleNumSemanasChange = (n) => {
    setNumSemanas(n);
    setCycle(generateCycle(catalog, n, seed));
  };

  const handleRegenerateCell = (semanaIdx, diaIdx, categoria) => {
    const pool = byCat[categoria];
    const current = cycle[semanaIdx][diaIdx][categoria];
    const candidates = pool.filter((d) => d.platillo !== current.platillo);
    const pick = candidates[Math.floor(Math.random() * candidates.length)];
    const newCycle = cycle.map((semana) => semana.map((dia) => ({ ...dia })));
    newCycle[semanaIdx][diaIdx][categoria] = pick;
    setCycle(newCycle);
  };

  const handleSelectDish = (semanaIdx, diaIdx, categoria, dish) => {
    const newCycle = cycle.map((semana) => semana.map((dia) => ({ ...dia })));
    newCycle[semanaIdx][diaIdx][categoria] = dish;
    setCycle(newCycle);
    setEditing(null);
  };

  const handleExportExcel = async () => {
    setExporting(true);
    try {
      await exportToExcel(cycle, numSemanas, clients[activeClient]?.name || 'cliente');
    } finally {
      setExporting(false);
    }
  };

  const handleAddClient = async (name, newCatalog) => {
    try {
      const created = await createCliente(name);
      await bulkInsertPlatillos(created.id, newCatalog);
      setClients((prev) => ({
        ...prev,
        [created.id]: { name: created.nombre, catalog: newCatalog },
      }));
      setShowNewClient(false);
      setActiveClient(created.id);
      setCycle(generateCycle(newCatalog, numSemanas, seed));
    } catch (e) {
      setLoadError('No se pudo crear el cliente. Intenta de nuevo.');
    }
  };

  if (loadError) {
    return (
      <div
        style={{
          padding: '40px 20px',
          textAlign: 'center',
          color: 'var(--color-text-danger)',
          background: 'var(--color-background-danger)',
          borderRadius: 'var(--border-radius-lg)',
          fontSize: '14px',
        }}
      >
        {loadError}
      </div>
    );
  }

  if (!loaded) {
    return (
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: '60px',
          color: 'var(--color-text-secondary)',
          gap: '8px',
        }}
      >
        <Loader2 size={16} className="spin" />
        Cargando datos...
      </div>
    );
  }

  const hasClients = Object.keys(clients).length > 0;

  return (
    <div style={{ fontFamily: 'var(--font-sans)', padding: '0', maxWidth: '100%' }}>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        .spin { animation: spin 1s linear infinite; }

        .menu-toolbar {
          display: flex;
          flex-wrap: wrap;
          align-items: center;
          gap: 12px;
          transition: max-height 0.25s ease, opacity 0.2s ease, padding 0.25s ease, margin 0.25s ease;
          overflow: hidden;
        }
        .menu-toolbar-group {
          display: flex;
          align-items: center;
          gap: 8px;
        }
        .menu-toolbar-actions {
          display: flex;
          gap: 8px;
          flex: 1 1 auto;
        }
        .menu-btn-label { display: inline; }

        .menu-toolbar-sticky {
          position: sticky;
          top: 0;
          z-index: 20;
          background: var(--color-background-primary);
          padding-top: 4px;
        }
        .menu-toolbar-mini {
          display: none;
        }
        .menu-toolbar-collapse-btn {
          display: none;
        }

        /* --- Tabla: columna de categoría fija fuera del scroll horizontal --- */
        .menu-scroll-wrap {
          position: relative;
          border: 0.5px solid var(--color-border-tertiary);
          border-radius: var(--border-radius-lg);
          overflow: hidden;
          background: var(--color-background-primary);
        }
        .menu-split {
          display: flex;
          align-items: flex-start;
        }
        .menu-label-col {
          flex-shrink: 0;
          display: flex;
          flex-direction: column;
          border-right: 1px solid var(--color-border-tertiary);
          z-index: 2;
        }
        .menu-days-scroll {
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
          scrollbar-width: thin;
          flex: 1;
          min-width: 0;
        }
        .menu-days-grid {
          display: grid;
        }
        .menu-cell {
          box-sizing: border-box;
          background: var(--color-background-primary);
          border-bottom: 0.5px solid var(--color-border-tertiary);
          border-right: 0.5px solid var(--color-border-tertiary);
          padding: 4px;
          display: flex;
          align-items: stretch;
          overflow: hidden;
          min-width: 0;
        }
        .menu-cell-corner,
        .menu-cell-daylabel {
          background: var(--color-background-info);
          color: var(--color-text-info);
          font-size: 11px;
          font-weight: 500;
          text-transform: uppercase;
          letter-spacing: 0.02em;
          padding: 6px;
          align-items: center;
        }
        .menu-cell-daylabel {
          justify-content: center;
          text-align: center;
          border-right: 0.5px solid rgba(255,255,255,0.15);
        }
        .menu-cell-corner {
          justify-content: flex-start;
          border-right: none;
        }
        .menu-cell-label {
          width: 100%;
          background: var(--color-background-secondary);
          color: var(--color-text-secondary);
          font-size: 11px;
          font-weight: 500;
          padding: 6px;
          align-items: center;
          border-right: none;
        }
        .menu-cell-dish {
          padding: 3px;
        }
        .menu-scroll-fade {
          position: absolute;
          top: 0;
          right: 0;
          bottom: 0;
          width: 16px;
          pointer-events: none;
          background: linear-gradient(to right, transparent, var(--color-background-primary) 95%);
          z-index: 3;
        }

        .menu-dish-cell-wrap {
          transition: border-color 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: flex-start;
          width: 100%;
        }
        .menu-dish-text {
          flex: 1;
          overflow: hidden;
        }
        .menu-dish-text-clamp {
          display: -webkit-box;
          -webkit-line-clamp: 2;
          -webkit-box-orient: vertical;
          overflow: hidden;
        }
        .menu-dish-actions-row {
          display: flex;
          justify-content: flex-end;
          gap: 3px;
          flex-shrink: 0;
          height: 18px;
          margin-bottom: 2px;
          opacity: 0.5;
          transition: opacity 0.15s ease;
        }
        .menu-dish-cell-wrap:hover .menu-dish-actions-row,
        .menu-dish-actions-row:focus-within {
          opacity: 1;
        }
        .menu-dish-actions-row button:hover {
          background: rgba(0,0,0,0.16) !important;
        }
        .menu-conflict-icon {
          margin-left: 4px;
          vertical-align: -1px;
          color: #BA7517;
          flex-shrink: 0;
        }

        .menu-info-backdrop {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(0,0,0,0.6);
          display: flex;
          align-items: center;
          justify-content: center;
          z-index: 999;
          padding: 24px;
          box-sizing: border-box;
        }
        .menu-info-box {
          background: var(--color-background-primary);
          color: var(--color-text-primary);
          border-radius: var(--border-radius-lg);
          padding: 18px 20px;
          width: 100%;
          max-width: 340px;
          max-height: 80vh;
          overflow-y: auto;
          box-sizing: border-box;
          position: relative;
          z-index: 1000;
          box-shadow: 0 8px 32px rgba(0,0,0,0.4);
        }
        .menu-info-title {
          font-size: 15px;
          font-weight: 600;
          line-height: 1.4;
          margin-bottom: 10px;
        }
        .menu-info-subtitle {
          font-size: 13px;
          color: var(--color-text-secondary);
          line-height: 1.4;
          margin-bottom: 8px;
        }
        .menu-info-list {
          margin: 0 0 14px;
          padding-left: 18px;
          font-size: 13px;
          line-height: 1.6;
          color: var(--color-text-primary);
        }
        .menu-info-close {
          width: 100%;
          padding: 9px;
          border-radius: var(--border-radius-md);
          border: 1px solid var(--color-border-secondary);
          background: var(--color-background-secondary);
          color: var(--color-text-primary);
          font-size: 14px;
          cursor: pointer;
        }

        @media (max-width: 680px) {
          .menu-toolbar-mini {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px 14px;
            background: var(--color-background-secondary);
            border-radius: var(--border-radius-lg);
            margin-bottom: 8px;
            cursor: pointer;
            font-size: 13px;
            color: var(--color-text-secondary);
          }
          .menu-toolbar-mini-text {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          .menu-toolbar-collapse-btn {
            display: flex;
            color: var(--color-text-tertiary);
            margin-left: 8px;
          }
          .menu-toolbar.is-collapsed {
            max-height: 0;
            opacity: 0;
            padding-top: 0;
            padding-bottom: 0;
            margin: 0;
            pointer-events: none;
            border-width: 0;
          }
          .menu-toolbar-group {
            justify-content: space-between;
          }
          .menu-toolbar-actions {
            flex-direction: column;
          }
          .menu-toolbar-actions button {
            width: 100%;
            justify-content: center;
          }
          .menu-toolbar-meta {
            justify-content: space-between;
            margin-left: 0 !important;
            width: 100%;
          }
          .menu-cell-corner,
          .menu-cell-daylabel { font-size: 11px !important; padding: 8px 4px !important; }
          .menu-cell-label { font-size: 11px !important; padding: 6px 4px !important; }
          .menu-dish-cell-wrap { font-size: 12px !important; }
          .menu-dish-actions-row { opacity: 1 !important; }
          .menu-scroll-fade {
            display: block;
            position: absolute;
            top: 0;
            right: 0;
            bottom: 0;
            width: 18px;
            pointer-events: none;
            background: linear-gradient(to right, transparent, var(--color-background-primary) 95%);
            z-index: 7;
          }
        }
      `}</style>
      <Toolbar
        clients={clients}
        activeClient={activeClient}
        onClientChange={handleClientChange}
        onAddClientClick={() => setShowNewClient(true)}
        numSemanas={numSemanas}
        onNumSemanasChange={handleNumSemanasChange}
        onRegenerate={handleRegenerate}
        onExport={handleExportExcel}
        exporting={exporting}
        catalogSize={catalog.length}
        saveStatus={saveStatus}
        collapsed={toolbarCollapsed}
        onToggleCollapsed={() => setToolbarCollapsed((v) => !v)}
      />
      {!hasClients ? (
        <EmptyState onAddClientClick={() => setShowNewClient(true)} />
      ) : catalog.length === 0 ? (
        <EmptyClientState clientName={clients[activeClient]?.name} />
      ) : (
        <>
          <Legend />
          {cycle &&
            cycle.map((semana, semanaIdx) => (
              <WeekTable
                key={semanaIdx}
                semanaIdx={semanaIdx}
                semana={semana}
                highlightWeek={highlights[semanaIdx]}
                onRegenerateCell={handleRegenerateCell}
                onOpenEdit={(diaIdx, categoria) =>
                  setEditing({ semana: semanaIdx, dia: diaIdx, categoria })
                }
                onOpenInfo={(platillo, conflicts) => setInfoModal({ platillo, conflicts })}
              />
            ))}
        </>
      )}
      {editing && (
        <EditModal
          categoria={editing.categoria}
          options={byCat[editing.categoria]}
          current={cycle[editing.semana][editing.dia][editing.categoria]}
          onSelect={(dish) =>
            handleSelectDish(editing.semana, editing.dia, editing.categoria, dish)
          }
          onClose={() => setEditing(null)}
        />
      )}
      {infoModal && (
        <InfoModal
          platillo={infoModal.platillo}
          conflicts={infoModal.conflicts}
          onClose={() => setInfoModal(null)}
        />
      )}
      {showNewClient && (
        <NewClientModal
          onClose={() => setShowNewClient(false)}
          onCreate={handleAddClient}
        />
      )}
    </div>
  );
}

function EmptyState({ onAddClientClick }) {
  return (
    <div
      style={{
        padding: '60px 20px',
        textAlign: 'center',
        color: 'var(--color-text-secondary)',
        background: 'var(--color-background-secondary)',
        borderRadius: 'var(--border-radius-lg)',
        fontSize: '14px',
      }}
    >
      Todavía no hay clientes registrados.
      <div style={{ marginTop: '12px' }}>
        <button onClick={onAddClientClick} style={btnStyle('primary')}>
          <Plus size={15} /> Agregar el primer cliente
        </button>
      </div>
    </div>
  );
}

function Toolbar({
  clients,
  activeClient,
  onClientChange,
  onAddClientClick,
  numSemanas,
  onNumSemanasChange,
  onRegenerate,
  onExport,
  exporting,
  catalogSize,
  saveStatus,
  collapsed,
  onToggleCollapsed,
}) {
  const activeName = clients[activeClient]?.name || '';
  return (
    <div className="menu-toolbar-sticky">
      {collapsed && (
        <div
          className="menu-toolbar-mini"
          onClick={onToggleCollapsed}
          role="button"
          tabIndex={0}
          aria-expanded={!collapsed}
        >
          <span className="menu-toolbar-mini-text">
            {activeName} · {numSemanas} {numSemanas === 1 ? 'semana' : 'semanas'}
          </span>
          <span className="menu-toolbar-collapse-btn" aria-hidden="true">
            ▾
          </span>
        </div>
      )}
      <div
        className={`menu-toolbar${collapsed ? ' is-collapsed' : ''}`}
        style={{
          padding: '16px',
          background: 'var(--color-background-secondary)',
          borderRadius: 'var(--border-radius-lg)',
          marginBottom: '16px',
        }}
      >
        <div className="menu-toolbar-group">
          <label style={{ fontSize: '14px', color: 'var(--color-text-secondary)' }}>
            Cliente:
          </label>
          <select
            value={activeClient}
            onChange={(e) => onClientChange(e.target.value)}
            style={{ ...selectStyle, flex: 1 }}
          >
            {Object.entries(clients).map(([id, c]) => (
              <option key={id} value={id}>
                {c.name}
              </option>
            ))}
          </select>
          <button onClick={onAddClientClick} title="Agregar cliente" style={iconBtnLarge}>
            <Plus size={15} />
          </button>
        </div>

        <div className="menu-toolbar-group">
          <label style={{ fontSize: '14px', color: 'var(--color-text-secondary)' }}>
            Semanas:
          </label>
          <select
            value={numSemanas}
            onChange={(e) => onNumSemanasChange(Number(e.target.value))}
            style={selectStyle}
          >
            {[1, 2, 3, 4, 6, 8, 10, 12].map((n) => (
              <option key={n} value={n}>
                {n}
              </option>
            ))}
          </select>
        </div>

        <div className="menu-toolbar-actions">
          <button onClick={onRegenerate} disabled={catalogSize === 0} style={btnStyle('primary')}>
            <RefreshCw size={15} /> <span className="menu-btn-label">Generar ciclo nuevo</span>
          </button>

          <button onClick={onExport} disabled={catalogSize === 0 || exporting} style={btnStyle('secondary')}>
            {exporting ? <Loader2 size={15} className="spin" /> : <Download size={15} />}{' '}
            <span className="menu-btn-label">{exporting ? 'Generando...' : 'Exportar a Excel'}</span>
          </button>
        </div>

        <div
          className="menu-toolbar-meta"
          style={{
            marginLeft: 'auto',
            fontSize: '12px',
            color: 'var(--color-text-tertiary)',
            display: 'flex',
            alignItems: 'center',
            gap: '12px',
          }}
        >
          <span>{catalogSize} platillos en catálogo</span>
          {saveStatus === 'saving' && (
            <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
              <Loader2 size={12} className="spin" /> Guardando...
            </span>
          )}
          {saveStatus === 'saved' && (
            <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
              <Check size={12} /> Guardado
            </span>
          )}
        </div>
      </div>
    </div>
  );
}

const selectStyle = {
  padding: '6px 10px',
  borderRadius: 'var(--border-radius-md)',
  border: '1px solid var(--color-border-secondary)',
  background: 'var(--color-background-primary)',
  color: 'var(--color-text-primary)',
  fontSize: '16px',
};

function btnStyle(variant) {
  const base = {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '6px',
    padding: '8px 14px',
    borderRadius: 'var(--border-radius-md)',
    fontSize: '14px',
    fontWeight: 500,
    cursor: 'pointer',
    border: 'none',
  };
  if (variant === 'primary') {
    return { ...base, background: '#1D9E75', color: '#04342C' };
  }
  return {
    ...base,
    background: 'var(--color-background-primary)',
    color: 'var(--color-text-primary)',
    border: '1px solid var(--color-border-secondary)',
  };
}

const iconBtnLarge = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: '30px',
  height: '30px',
  borderRadius: 'var(--border-radius-md)',
  border: '1px solid var(--color-border-secondary)',
  background: 'var(--color-background-primary)',
  color: 'var(--color-text-primary)',
  cursor: 'pointer',
};

function Legend() {
  return (
    <div
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: '8px',
        fontSize: '13px',
        color: 'var(--color-text-secondary)',
        marginBottom: '16px',
        flexWrap: 'wrap',
      }}
    >
      <AlertTriangle size={14} style={{ color: '#BA7517' }} />
      <span>
        Las celdas resaltadas se repiten con otra dentro de {COOLDOWN_WEEKS} semanas — tócalas para ver con cuál.
      </span>
    </div>
  );
}

function EmptyClientState({ clientName }) {
  return (
    <div
      style={{
        padding: '60px 20px',
        textAlign: 'center',
        color: 'var(--color-text-secondary)',
        background: 'var(--color-background-secondary)',
        borderRadius: 'var(--border-radius-lg)',
        fontSize: '14px',
      }}
    >
      {clientName} todavía no tiene catálogo de platillos cargado.
    </div>
  );
}

const LABEL_COL_WIDTH = 100; // px, ancho de la columna de categoría (fuera del scroll horizontal)
const DAY_COL_WIDTH = 116; // px, ancho de cada columna de día (dentro del scroll horizontal)
const HEADER_ROW_H = 36; // px, alto de la fila de encabezado (días)
const BODY_ROW_H = 86; // px, alto de cada fila de categoría

function WeekTable({ semanaIdx, semana, highlightWeek, onRegenerateCell, onOpenEdit, onOpenInfo }) {
  return (
    <div style={{ marginBottom: '24px' }}>
      <div
        style={{
          fontSize: '16px',
          fontWeight: 500,
          marginBottom: '8px',
          color: 'var(--color-text-primary)',
        }}
      >
        Semana {semanaIdx + 1}
      </div>
      <div className="menu-scroll-wrap">
        <div className="menu-split">
          {/* Columna de categoría: nunca se mueve, no forma parte del scroll horizontal */}
          <div
            className="menu-label-col"
            style={{ width: LABEL_COL_WIDTH }}
            role="rowgroup"
            aria-label="Categorías"
          >
            <div className="menu-cell menu-cell-corner" style={{ height: HEADER_ROW_H }} role="columnheader">
              Categoría
            </div>
            {CATEGORIAS.map((cat) => (
              <div
                key={cat}
                className="menu-cell menu-cell-label"
                style={{ height: BODY_ROW_H }}
                role="rowheader"
              >
                {CATEGORIA_LABEL_TABLA[cat]}
              </div>
            ))}
          </div>

          {/* Días + platillos: única zona con scroll horizontal */}
          <div className="menu-days-scroll">
            <div
              className="menu-days-grid"
              style={{
                gridTemplateColumns: `repeat(${DIAS.length}, minmax(${DAY_COL_WIDTH}px, 1fr))`,
                minWidth: DAY_COL_WIDTH * DIAS.length,
              }}
              role="table"
              aria-label={`Menú semana ${semanaIdx + 1}`}
            >
              {DIAS.map((dia) => (
                <div
                  key={dia}
                  className="menu-cell menu-cell-daylabel"
                  style={{ height: HEADER_ROW_H }}
                  role="columnheader"
                >
                  {dia}
                </div>
              ))}
              {CATEGORIAS.map((cat) =>
                DIAS.map((_, diaIdx) => {
                  const dish = semana[diaIdx][cat];
                  const highlight = highlightWeek[diaIdx][cat];
                  return (
                    <div
                      key={`${cat}_${diaIdx}`}
                      className="menu-cell menu-cell-dish"
                      style={{ height: BODY_ROW_H }}
                      role="cell"
                    >
                      <DishCell
                        dish={dish}
                        highlight={highlight}
                        isFixed={!!FIXED_CATEGORIES[cat]}
                        onRegenerate={() => onRegenerateCell(semanaIdx, diaIdx, cat)}
                        onEdit={() => onOpenEdit(diaIdx, cat)}
                        onOpenInfo={onOpenInfo}
                      />
                    </div>
                  );
                })
              )}
            </div>
          </div>
        </div>
        <div className="menu-scroll-fade" aria-hidden="true" />
      </div>
    </div>
  );
}

function headerCellStyle(align) {
  return {
    fontSize: '12px',
    fontWeight: 500,
    textAlign: align,
    padding: '10px 12px',
    textTransform: 'uppercase',
    letterSpacing: '0.02em',
  };
}

const rowLabelStyle = {
  fontSize: '13px',
  fontWeight: 500,
  color: 'var(--color-text-secondary)',
  padding: '10px 12px',
  whiteSpace: 'normal',
  verticalAlign: 'middle',
  wordBreak: 'break-word',
};

function DishCell({ dish, highlight, isFixed, onRegenerate, onEdit, onOpenInfo }) {
  if (!dish) return null;

  const hasConflict = !!highlight;
  const style = {
    background: hasConflict ? CONFLICT_STYLE.bg : 'var(--color-background-primary)',
    border: `1px solid ${hasConflict ? CONFLICT_STYLE.border : 'var(--color-border-tertiary)'}`,
    borderRadius: 'var(--border-radius-md)',
    padding: isFixed ? '8px 10px' : '4px 8px 8px',
    height: '100%',
    boxSizing: 'border-box',
    fontSize: '12px',
    color: hasConflict ? CONFLICT_STYLE.text : 'var(--color-text-primary)',
    lineHeight: 1.3,
  };

  const isLong = dish.platillo.length > 30;
  const isTappable = isLong || hasConflict;

  const handleTap = isTappable
    ? (e) => {
        e.stopPropagation();
        onOpenInfo(dish.platillo, hasConflict ? highlight.conflicts : null);
      }
    : undefined;

  return (
    <div className="menu-dish-cell-wrap" style={style}>
      {!isFixed && (
        <div className="menu-dish-actions-row">
          <button
            onClick={(e) => { e.stopPropagation(); onRegenerate(); }}
            title="Sustituir aleatoriamente"
            aria-label="Sustituir platillo aleatoriamente"
            style={iconBtnStyle}
          >
            <Shuffle size={11} />
          </button>
          <button
            onClick={(e) => { e.stopPropagation(); onEdit(); }}
            title="Elegir manualmente"
            aria-label="Elegir platillo manualmente"
            style={iconBtnStyle}
          >
            <Edit3 size={11} />
          </button>
        </div>
      )}
      <div
        className={isLong ? 'menu-dish-text menu-dish-text-clamp' : 'menu-dish-text'}
        style={{ cursor: isTappable ? 'pointer' : 'default' }}
        onClick={handleTap}
        title={isLong ? dish.platillo : undefined}
      >
        {dish.platillo}
        {hasConflict && <AlertTriangle size={11} className="menu-conflict-icon" aria-label="Se repite" />}
      </div>
    </div>
  );
}

const iconBtnStyle = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: '20px',
  height: '20px',
  borderRadius: '5px',
  border: 'none',
  background: 'rgba(0,0,0,0.08)',
  color: 'inherit',
  cursor: 'pointer',
};

function InfoModal({ platillo, conflicts, onClose }) {
  return createPortal(
    <div className="menu-info-backdrop" onClick={onClose}>
      <div className="menu-info-box" onClick={(e) => e.stopPropagation()}>
        <div className="menu-info-title">{platillo}</div>
        {conflicts && conflicts.length > 0 && (
          <>
            <div className="menu-info-subtitle">
              Se repite (misma familia) en {conflicts.length === 1 ? 'esta otra fecha' : 'estas otras fechas'}, dentro de las {COOLDOWN_WEEKS} semanas de margen:
            </div>
            <ul className="menu-info-list">
              {conflicts.map((c, i) => (
                <li key={i}>
                  <strong>Semana {c.semanaIdx + 1}, {DIAS[c.diaIdx]}:</strong> {c.platillo}
                </li>
              ))}
            </ul>
          </>
        )}
        <button className="menu-info-close" onClick={onClose}>
          Cerrar
        </button>
      </div>
    </div>,
    document.body
  );
}

function EditModal({ categoria, options, current, onSelect, onClose }) {
  const [query, setQuery] = useState('');
  const filtered = options.filter((d) =>
    d.platillo.toLowerCase().includes(query.toLowerCase())
  );

  return createPortal(
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.45)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 50,
      }}
      onClick={onClose}
    >
      <div
        style={{
          background: 'var(--color-background-primary)',
          borderRadius: 'var(--border-radius-lg)',
          padding: '20px',
          width: '420px',
          maxWidth: 'calc(100vw - 32px)',
          maxHeight: '70vh',
          display: 'flex',
          flexDirection: 'column',
          boxSizing: 'border-box',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: '12px',
          }}
        >
          <div style={{ fontSize: '16px', fontWeight: 500 }}>
            {CATEGORIA_LABEL[categoria]}
          </div>
          <button onClick={onClose} style={{ ...iconBtnStyle, width: '24px', height: '24px' }}>
            <X size={14} />
          </button>
        </div>
        <input
          autoFocus
          placeholder="Buscar platillo..."
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          style={{
            padding: '10px',
            borderRadius: 'var(--border-radius-md)',
            border: '1px solid var(--color-border-secondary)',
            background: 'var(--color-background-primary)',
            color: 'var(--color-text-primary)',
            fontSize: '16px',
            marginBottom: '10px',
            boxSizing: 'border-box',
            width: '100%',
          }}
        />
        <div style={{ overflowY: 'auto', flex: 1 }}>
          {filtered.map((d) => (
            <div
              key={d.id}
              onClick={() => onSelect(d)}
              style={{
                padding: '8px 10px',
                borderRadius: 'var(--border-radius-md)',
                cursor: 'pointer',
                fontSize: '13px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                background:
                  d.platillo === current.platillo
                    ? 'var(--color-background-secondary)'
                    : 'transparent',
              }}
              onMouseEnter={(e) =>
                (e.currentTarget.style.background = 'var(--color-background-secondary)')
              }
              onMouseLeave={(e) =>
                (e.currentTarget.style.background =
                  d.platillo === current.platillo
                    ? 'var(--color-background-secondary)'
                    : 'transparent')
              }
            >
              <span>{d.platillo}</span>
              {d.platillo === current.platillo && <Check size={14} />}
            </div>
          ))}
        </div>
      </div>
    </div>,
    document.body
  );
}

function NewClientModal({ onClose, onCreate }) {
  const [name, setName] = useState('');
  const [fileData, setFileData] = useState(null);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setError('');
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const sheetName =
        wb.SheetNames.find((n) => /consolidad/i.test(n)) || wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      // Espera columnas: TIEMPO, PLATILLO, TIPO, INGREDIENTES, ALERGENO, KCAL
      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || !row[1]) continue;
        parsed.push({
          categoria: String(row[0] || '').trim().toUpperCase(),
          platillo: String(row[1]).trim(),
          tipo: String(row[2] || '').trim(),
          ingredientes: String(row[3] || '').trim(),
          alergeno: String(row[4] || '').trim(),
          kcal: row[5] || '',
        });
      }
      if (parsed.length === 0) {
        setError('No se encontraron platillos en la hoja. Revisa el formato.');
        return;
      }
      // id + familia_id simple (por nombre exacto normalizado) para datos nuevos
      const seen = {};
      let famCounter = 0;
      parsed.forEach((d, idx) => {
        d.id = `new_${idx}`;
        const norm = d.platillo.toLowerCase().trim();
        if (!(norm in seen)) {
          seen[norm] = `fam_new_${famCounter++}`;
        }
        d.familia_id = seen[norm];
        d.keywords = [];
      });
      setFileData(parsed);
    } catch (err) {
      setError('No se pudo leer el archivo. Verifica que sea un .xlsx válido.');
    }
  };

  return createPortal(
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.45)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 50,
      }}
      onClick={onClose}
    >
      <div
        style={{
          background: 'var(--color-background-primary)',
          borderRadius: 'var(--border-radius-lg)',
          padding: '24px',
          width: '420px',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: '16px',
          }}
        >
          <div style={{ fontSize: '16px', fontWeight: 500 }}>Nuevo cliente</div>
          <button onClick={onClose} style={{ ...iconBtnStyle, width: '24px', height: '24px' }}>
            <X size={14} />
          </button>
        </div>

        <label style={{ fontSize: '13px', color: 'var(--color-text-secondary)' }}>
          Nombre del cliente
        </label>
        <input
          autoFocus
          placeholder="Ej. Netflix"
          value={name}
          onChange={(e) => setName(e.target.value)}
          style={{
            width: '100%',
            padding: '10px',
            borderRadius: 'var(--border-radius-md)',
            border: '1px solid var(--color-border-secondary)',
            background: 'var(--color-background-primary)',
            color: 'var(--color-text-primary)',
            fontSize: '16px',
            marginTop: '6px',
            marginBottom: '16px',
            boxSizing: 'border-box',
          }}
        />

        <label style={{ fontSize: '13px', color: 'var(--color-text-secondary)' }}>
          Excel del consolidado de platillos
        </label>
        <input
          type="file"
          accept=".xlsx"
          onChange={handleFile}
          style={{ display: 'block', marginTop: '6px', marginBottom: '8px', fontSize: '13px' }}
        />
        {fileName && !error && (
          <div style={{ fontSize: '12px', color: 'var(--color-text-tertiary)' }}>
            {fileData ? `${fileData.length} platillos detectados en "${fileName}"` : 'Procesando...'}
          </div>
        )}
        {error && (
          <div style={{ fontSize: '12px', color: '#A32D2D' }}>{error}</div>
        )}

        <button
          onClick={() => onCreate(name.trim(), fileData)}
          disabled={!name.trim() || !fileData}
          style={{
            ...btnStyle('primary'),
            width: '100%',
            justifyContent: 'center',
            marginTop: '18px',
            opacity: !name.trim() || !fileData ? 0.5 : 1,
          }}
        >
          Crear cliente
        </button>
      </div>
    </div>,
    document.body
  );
}

async function exportToExcel(cycle, numSemanas, clientName) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Ciclo de menus');

  const HEADER_BLUE = 'FF29ABE2';
  const LABEL_GRAY = 'FFE8E8E8';
  const BORDER = { style: 'thin', color: { argb: 'FFBFBFBF' } };
  const thinBorder = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
  const colsPerWeek = 6; // 5 días + 1 columna separadora

  // Fila 1: "SEMANA 01" combinada sobre sus 5 días
  let col = 2;
  cycle.forEach((_, sIdx) => {
    const startCol = col;
    const endCol = col + 4;
    ws.mergeCells(1, startCol, 1, endCol);
    const cell = ws.getCell(1, startCol);
    cell.value = `SEMANA ${String(sIdx + 1).padStart(2, '0')}`;
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BLUE } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    col += colsPerWeek;
  });

  // Fila 2: nombre de cada día + "DESAYUNO" en A2
  const labelHeaderCell = ws.getCell(2, 1);
  labelHeaderCell.value = 'DESAYUNO';
  labelHeaderCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BLUE } };
  labelHeaderCell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  labelHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };

  col = 2;
  cycle.forEach(() => {
    DIAS.forEach((dia) => {
      const cell = ws.getCell(2, col);
      cell.value = dia;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BLUE } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      col++;
    });
    col++; // separador
  });

  // Filas de categorías
  let rowIdx = 3;
  CATEGORIAS.forEach((cat) => {
    const labelCell = ws.getCell(rowIdx, 1);
    labelCell.value = CATEGORIA_LABEL[cat];
    labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LABEL_GRAY } };
    labelCell.font = { bold: true, size: 10 };
    labelCell.alignment = { vertical: 'middle', wrapText: true };
    labelCell.border = thinBorder;

    let c = 2;
    cycle.forEach((semana) => {
      DIAS.forEach((_, dIdx) => {
        const cell = ws.getCell(rowIdx, c);
        cell.value = semana[dIdx][cat]?.platillo || '';
        cell.alignment = { vertical: 'middle', wrapText: true };
        cell.border = thinBorder;
        cell.font = { size: 10 };
        c++;
      });
      c++; // separador
    });
    rowIdx++;
  });

  // Anchos de columna
  ws.getColumn(1).width = 28;
  let c = 2;
  cycle.forEach(() => {
    for (let i = 0; i < 5; i++) {
      ws.getColumn(c).width = 22;
      c++;
    }
    ws.getColumn(c).width = 3; // separador angosto
    c++;
  });

  // Alto de filas para que el texto envuelto se vea completo
  for (let r = 3; r < rowIdx; r++) {
    ws.getRow(r).height = 45;
  }

  // Congela encabezados (filas 1-2) y la columna de categorías
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }];

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const safeClient = (clientName || 'cliente').replace(/[^a-z0-9]+/gi, '_');
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `ciclo_menus_${safeClient}_${numSemanas}semanas.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}