import { createClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabase = createClient(supabaseUrl, supabaseKey);

// --- Clientes -------------------------------------------------------

export async function listClientes() {
  const { data, error } = await supabase.from('clientes').select('*').order('nombre');
  if (error) throw error;
  return data;
}

export async function createCliente(nombre) {
  const slug = nombre.toLowerCase().trim().replace(/[^a-z0-9]+/g, '_');
  const { data, error } = await supabase
    .from('clientes')
    .insert({ nombre, slug })
    .select()
    .single();
  if (error) throw error;
  return data;
}

// --- Platillos --------------------------------------------------------

export async function listPlatillos(clienteId) {
  const { data, error } = await supabase
    .from('platillos')
    .select('*')
    .eq('cliente_id', clienteId);
  if (error) throw error;
  // adaptar nombres de columna (snake_case en BD) al formato que usa el componente
  return data.map((d) => ({
    id: d.id,
    categoria: d.categoria,
    platillo: d.platillo,
    tipo: d.tipo,
    ingredientes: d.ingredientes,
    alergeno: d.alergeno,
    kcal: d.kcal,
    familia_id: d.familia_id,
  }));
}

// Inserta el catálogo completo de un cliente de una sola vez.
// dishes: array de { categoria, platillo, tipo, ingredientes, alergeno, kcal, familia_id }
export async function bulkInsertPlatillos(clienteId, dishes) {
  const rows = dishes.map((d) => ({
    cliente_id: clienteId,
    categoria: d.categoria,
    platillo: d.platillo,
    tipo: d.tipo || '',
    ingredientes: d.ingredientes || '',
    alergeno: d.alergeno || '',
    kcal: d.kcal ? String(d.kcal) : '',
    familia_id: d.familia_id,
  }));
  // Supabase recomienda insertar en lotes de ~500 para evitar payloads gigantes
  const BATCH = 400;
  for (let i = 0; i < rows.length; i += BATCH) {
    const batch = rows.slice(i, i + BATCH);
    const { error } = await supabase.from('platillos').insert(batch);
    if (error) throw error;
  }
}

// --- Ciclos generados ---------------------------------------------------

export async function getLastCiclo(clienteId) {
  const { data, error } = await supabase
    .from('ciclos')
    .select('*')
    .eq('cliente_id', clienteId)
    .order('created_at', { ascending: false })
    .limit(1)
    .maybeSingle();
  if (error) throw error;
  return data;
}

export async function saveCiclo(clienteId, numSemanas, seed, cycleData) {
  const { error } = await supabase.from('ciclos').insert({
    cliente_id: clienteId,
    num_semanas: numSemanas,
    seed,
    datos: cycleData,
  });
  if (error) throw error;
}
