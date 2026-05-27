// lib/supabase.ts
import { createClient } from '@supabase/supabase-js';

const supabaseUrl  = process.env.NEXT_PUBLIC_SUPABASE_URL!;
const supabaseAnon = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!;

export const supabase = createClient(supabaseUrl, supabaseAnon, {
  auth: {
    persistSession:    true,
    autoRefreshToken:  true,   // FIX ERROR 2: sesión nunca expira mientras usa la app
    detectSessionInUrl: false,
  },
});

export type UserRole = 'admin' | 'lider' | 'tecnico' | 'gerencia' | 'cliente' | 'visualizador';

export interface Profile {
  id:              string;
  nombre:          string;
  correo:          string;
  rol:             UserRole;
  especialidad_id?: string;
  activo:          boolean;
}
