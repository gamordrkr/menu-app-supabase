import React, { useState, useEffect, useMemo, useCallback } from 'react';
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

const DIAS = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'];

const FAMILY_COLORS = [
  { bg: '#FAEEDA', border: '#BA7517', text: '#633806' },
  { bg: '#EAF3DE', border: '#3B6D11', text: '#27500A' },
  { bg: '#FBEAF0', border: '#993556', text: '#72243E' },
  { bg: '#E6F1FB', border: '#185FA5', text: '#0C447C' },
  { bg: '#FCEBEB', border: '#A32D2D', text: '#791F1F' },
  { bg: '#EEEDFE', border: '#534AB7', text: '#3C3489' },
];

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

// Resalta SOLO familias que quedaron a menos de COOLDOWN_WEEKS de distancia
// real (incluye casos donde, por falta de variedad, el algoritmo tuvo que
// relajar la regla) para que el usuario las revise.
function computeViolationHighlights(cycle) {
  const lastSeenWeek = {};
  for (const cat of CATEGORIAS) lastSeenWeek[cat] = {};
  const highlights = []; // [semanaIdx][diaIdx][categoria] = colorObj | null
  let colorCursor = 0;
  const familyColorAssign = {};

  for (let s = 0; s < cycle.length; s++) {
    const semanaHighlight = DIAS.map(() => ({}));
    for (const cat of CATEGORIAS) {
      if (FIXED_CATEGORIES[cat]) {
        DIAS.forEach((_, d) => (semanaHighlight[d][cat] = null));
        continue;
      }
      for (let d = 0; d < 5; d++) {
        const dish = cycle[s][d][cat];
        if (!dish) continue;
        const lastWeek = lastSeenWeek[cat][dish.familia_id];
        const isViolation = lastWeek !== undefined && s - lastWeek < COOLDOWN_WEEKS;
        if (isViolation) {
          const key = `${cat}_${dish.familia_id}`;
          if (!familyColorAssign[key]) {
            familyColorAssign[key] = FAMILY_COLORS[colorCursor % FAMILY_COLORS.length];
            colorCursor++;
          }
          semanaHighlight[d][cat] = familyColorAssign[key];
        } else {
          semanaHighlight[d][cat] = null;
        }
      }
    }
    // actualizar lastSeenWeek con lo que se vio esta semana
    for (const cat of CATEGORIAS) {
      if (FIXED_CATEGORIES[cat]) continue;
      for (let d = 0; d < 5; d++) {
        const dish = cycle[s][d][cat];
        if (dish) lastSeenWeek[cat][dish.familia_id] = s;
      }
    }
    highlights.push(semanaHighlight);
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
  const [loaded, setLoaded] = useState(false);
  const [showNewClient, setShowNewClient] = useState(false);
  const [saveStatus, setSaveStatus] = useState('idle'); // idle | saving | saved | error
  const [loadError, setLoadError] = useState('');
  const [exporting, setExporting] = useState(false);
  const catalog = clients[activeClient]?.catalog || [];
  const byCat = useMemo(() => buildIndexByCategory(catalog), [catalog]);

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

        /* --- Tabla con encabezado y primera columna fijos (sticky) --- */
        .menu-scroll-wrap {
          position: relative;
          border: 0.5px solid var(--color-border-tertiary);
          border-radius: var(--border-radius-lg);
          overflow: hidden;
          background: var(--color-background-primary);
        }
        .menu-scroll-area {
          overflow: auto;
          -webkit-overflow-scrolling: touch;
          scrollbar-width: thin;
          max-height: 70vh;
        }
        .menu-table {
          border-collapse: separate;
          border-spacing: 0;
          width: 100%;
          background: var(--color-background-primary);
          box-sizing: border-box;
        }
        .menu-col-label { width: 18%; }
        .menu-col-day { width: 16.4%; }

        .menu-table th,
        .menu-table td {
          background: var(--color-background-primary);
          box-sizing: border-box;
          overflow: hidden;
        }
        .menu-table thead th {
          position: sticky;
          top: 0;
          z-index: 3;
          background: var(--color-background-info) !important;
          color: var(--color-text-info);
          box-shadow: 0 1px 0 0 var(--color-border-tertiary);
        }
        .menu-row-label {
          position: sticky;
          left: 0;
          z-index: 2;
          background: var(--color-background-secondary) !important;
          box-shadow: 1px 0 0 0 var(--color-border-tertiary);
        }
        .menu-table thead th.menu-corner {
          left: 0;
          z-index: 4;
          background: var(--color-background-info) !important;
          box-shadow: 1px 1px 0 0 var(--color-border-tertiary);
        }
        .menu-dish-td {
          padding: 4px;
          vertical-align: top;
        }
        .menu-scroll-fade {
          display: none;
        }

        .menu-dish-cell-wrap {
          transition: border-color 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: flex-start;
          height: 100%;
        }
        .menu-dish-text {
          flex: 1;
        }
        .menu-dish-text-clamp {
          display: -webkit-box;
          -webkit-line-clamp: 3;
          -webkit-box-orient: vertical;
          overflow: hidden;
        }
        .menu-dish-actions {
          position: absolute;
          top: 6px;
          right: 6px;
          display: flex;
          gap: 4px;
          opacity: 0.55;
          transition: opacity 0.15s ease;
        }
        .menu-dish-cell-wrap:hover .menu-dish-actions,
        .menu-dish-actions:focus-within {
          opacity: 1;
        }
        .menu-dish-actions button:hover {
          background: rgba(0,0,0,0.16) !important;
        }

        @media (max-width: 680px) {
          .menu-toolbar {
            flex-direction: column;
            align-items: stretch;
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
          .menu-col-label { width: 96px; }
          .menu-col-day { width: 108px; }
          .menu-table { min-width: 0; }
          .menu-table th { font-size: 11px !important; padding: 8px 6px !important; }
          .menu-row-label { font-size: 11px !important; padding: 6px !important; }
          .menu-dish-cell-wrap { font-size: 12px !important; padding: 8px !important; min-height: 56px !important; }
          .menu-dish-actions { opacity: 1 !important; gap: 6px !important; }
          .menu-dish-actions button { width: 30px !important; height: 30px !important; }
          .menu-scroll-fade {
            display: block;
            position: absolute;
            top: 0;
            right: 0;
            bottom: 0;
            width: 18px;
            pointer-events: none;
            background: linear-gradient(to right, transparent, var(--color-background-primary) 95%);
            z-index: 4;
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
}) {
  return (
    <div
      className="menu-toolbar"
      style={{
        padding: '16px',
        background: 'var(--color-background-secondary)',
        borderRadius: 'var(--border-radius-lg)',
        marginBottom: '16px',
        position: 'sticky',
        top: 0,
        zIndex: 10,
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
          {[1, 2, 3, 4, 6, 8].map((n) => (
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
        Las celdas resaltadas comparten familia con otra aparecida hace menos de{' '}
        {COOLDOWN_WEEKS} semanas — revísalas antes de aprobar.
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

function WeekTable({ semanaIdx, semana, highlightWeek, onRegenerateCell, onOpenEdit }) {
  return (
    <div style={{ marginBottom: '28px' }}>
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
        <div className="menu-scroll-area">
          <table
            className="menu-table"
            style={{
              width: '100%',
              minWidth: '760px',
              tableLayout: 'fixed',
            }}
          >
            <colgroup>
              <col className="menu-col-label" />
              {DIAS.map((dia) => (
                <col key={dia} className="menu-col-day" />
              ))}
            </colgroup>
            <thead>
              <tr>
                <th className="menu-corner" style={headerCellStyle('left')}>
                  <span style={{ display: 'block', textAlign: 'left' }}>Categoría</span>
                </th>
                {DIAS.map((dia) => (
                  <th key={dia} style={headerCellStyle('center')}>
                    <span style={{ display: 'block', textAlign: 'center' }}>{dia}</span>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {CATEGORIAS.map((cat) => (
                <tr key={cat}>
                  <td className="menu-row-label" style={rowLabelStyle}>
                    {CATEGORIA_LABEL[cat]}
                  </td>
                  {DIAS.map((_, diaIdx) => {
                    const dish = semana[diaIdx][cat];
                    const highlight = highlightWeek[diaIdx][cat];
                    return (
                      <td key={diaIdx} className="menu-dish-td">
                        <DishCell
                          dish={dish}
                          highlight={highlight}
                          isFixed={!!FIXED_CATEGORIES[cat]}
                          onRegenerate={() => onRegenerateCell(semanaIdx, diaIdx, cat)}
                          onEdit={() => onOpenEdit(diaIdx, cat)}
                        />
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
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

function DishCell({ dish, highlight, isFixed, onRegenerate, onEdit }) {
  const [showFull, setShowFull] = useState(false);
  if (!dish) return null;

  const style = {
    position: 'relative',
    background: highlight ? highlight.bg : 'var(--color-background-primary)',
    border: `1px solid ${highlight ? highlight.border : 'var(--color-border-tertiary)'}`,
    borderRadius: 'var(--border-radius-md)',
    padding: '10px 12px',
    minHeight: '64px',
    height: '100%',
    boxSizing: 'border-box',
    fontSize: '13px',
    color: highlight ? highlight.text : 'var(--color-text-primary)',
    lineHeight: 1.4,
  };

  // Texto largo (más de ~38 caracteres) se trunca a 3 líneas; el resto se ve completo.
  const isLong = dish.platillo.length > 38;

  return (
    <div className="menu-dish-cell-wrap" style={style}>
      <div
        className={isLong ? 'menu-dish-text menu-dish-text-clamp' : 'menu-dish-text'}
        style={{ paddingRight: isFixed ? 0 : '52px', cursor: isLong ? 'pointer' : 'default' }}
        onClick={isLong ? (e) => { e.stopPropagation(); setShowFull(true); } : undefined}
        title={isLong ? dish.platillo : undefined}
      >
        {dish.platillo}
      </div>
      {!isFixed && (
        <div className="menu-dish-actions">
          <button
            onClick={(e) => {
              e.stopPropagation();
              onRegenerate();
            }}
            title="Sustituir aleatoriamente"
            aria-label="Sustituir platillo aleatoriamente"
            style={iconBtnStyle}
          >
            <Shuffle size={13} />
          </button>
          <button
            onClick={(e) => {
              e.stopPropagation();
              onEdit();
            }}
            title="Elegir manualmente"
            aria-label="Elegir platillo manualmente"
            style={iconBtnStyle}
          >
            <Edit3 size={13} />
          </button>
        </div>
      )}
      {showFull && (
        <div
          role="dialog"
          aria-label="Nombre completo del platillo"
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(0,0,0,0.45)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 60,
          }}
          onClick={(e) => {
            e.stopPropagation();
            setShowFull(false);
          }}
        >
          <div
            style={{
              background: 'var(--color-background-primary)',
              color: 'var(--color-text-primary)',
              borderRadius: 'var(--border-radius-lg)',
              padding: '18px 20px',
              maxWidth: 'calc(100vw - 48px)',
              width: '340px',
              fontSize: '14px',
              lineHeight: 1.5,
              boxSizing: 'border-box',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {dish.platillo}
          </div>
        </div>
      )}
    </div>
  );
}

const iconBtnStyle = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: '26px',
  height: '26px',
  borderRadius: '6px',
  border: 'none',
  background: 'rgba(0,0,0,0.08)',
  color: 'inherit',
  cursor: 'pointer',
};

function EditModal({ categoria, options, current, onSelect, onClose }) {
  const [query, setQuery] = useState('');
  const filtered = options.filter((d) =>
    d.platillo.toLowerCase().includes(query.toLowerCase())
  );

  return (
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
    </div>
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

  return (
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
    </div>
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