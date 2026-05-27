// app/api/admin/route.ts
// Operaciones admin que requieren service_role key (nunca va al browser)
import { NextRequest, NextResponse } from 'next/server';
import { createClient } from '@supabase/supabase-js';

function adminClient() {
  return createClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.SUPABASE_SERVICE_ROLE_KEY!,
    { auth: { autoRefreshToken: false, persistSession: false } }
  );
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { action } = body;

    if (action === 'createUser') {
      const { email, password, nombre, rol, especialidad_id } = body;
      if (!email || !password || !nombre || !rol) {
        return NextResponse.json({ ok: false, error: 'Faltan datos' }, { status: 400 });
      }
      const admin = adminClient();
      const { data, error } = await admin.auth.admin.createUser({
        email, password,
        email_confirm: true,
        user_metadata: { nombre, rol, especialidad_id },
      });
      if (error) return NextResponse.json({ ok: false, error: error.message }, { status: 400 });

      // Asegurar que el perfil tiene el rol correcto
      await admin.from('profiles').upsert({
        id: data.user.id, nombre, correo: email, rol,
        especialidad_id: especialidad_id || null, activo: true,
      });
      return NextResponse.json({ ok: true, user_id: data.user.id });
    }

    if (action === 'deleteUser') {
      const { user_id } = body;
      const admin = adminClient();
      const { error } = await admin.auth.admin.deleteUser(user_id);
      if (error) return NextResponse.json({ ok: false, error: error.message }, { status: 400 });
      return NextResponse.json({ ok: true });
    }

    if (action === 'updateUser') {
      const { user_id, nombre, rol, especialidad_id, activo, password } = body;
      const admin = adminClient();
      if (password) {
        await admin.auth.admin.updateUserById(user_id, { password });
      }
      await admin.from('profiles').update({
        nombre, rol, especialidad_id: especialidad_id || null, activo,
        updated_at: new Date().toISOString(),
      }).eq('id', user_id);
      return NextResponse.json({ ok: true });
    }

    if (action === 'listUsers') {
      const admin = adminClient();
      const { data, error } = await admin.from('profiles').select('*').order('nombre');
      if (error) return NextResponse.json({ ok: false, error: error.message }, { status: 400 });
      return NextResponse.json({ ok: true, users: data });
    }

    return NextResponse.json({ ok: false, error: 'Acción no reconocida' }, { status: 400 });
  } catch (err: any) {
    return NextResponse.json({ ok: false, error: err?.message || 'Error interno' }, { status: 500 });
  }
}
