# Checklist de instalación — PDS 360 v2.0 (Supabase)

---

## PASO 1 — Supabase: crear las tablas

- [ ] Abierto supabase.com con dcetinag@gmail.com
- [ ] Proyecto `pds-360` ya creado y activo
- [ ] Ir a **SQL Editor → New query**
- [ ] Copiado el contenido de `supabase/schema.sql`
- [ ] Ejecutado el SQL (botón Run o Ctrl+Enter)
- [ ] Aparece "Success" sin errores
- [ ] Verificado en **Table Editor** que hay 20+ tablas

---

## PASO 2 — Supabase: primer usuario admin

- [ ] Ir a **Authentication → Users → Add user**
- [ ] Email: `dcetinag@gmail.com` · Password: la que quieras
- [ ] Clic en **Create user**
- [ ] En SQL Editor ejecutar:
  ```sql
  update profiles set rol = 'admin', nombre = 'David Cetina'
  where correo = 'dcetinag@gmail.com';
  ```
- [ ] Verificado que el perfil existe en Table Editor → profiles

---

## PASO 3 — Proyecto local

- [ ] Carpeta `pds-360` creada con `create-next-app`
- [ ] `npm install @supabase/supabase-js xlsx` ejecutado
- [ ] Archivos de este ZIP copiados al proyecto
- [ ] `.env.local.example` copiado como `.env.local`
- [ ] `npm run dev` corre sin errores en http://localhost:3000

---

## PASO 4 — Prueba local

- [ ] Login con dcetinag@gmail.com funciona
- [ ] Dashboard carga sin errores
- [ ] Módulo Usuarios carga la lista

---

## PASO 5 — Cargar catálogos iniciales

- [ ] Especialidades y actividades cargadas (desde V1 o Excel nuevo)
- [ ] Áreas cargadas (PREVIO1, PREVIO2, P1-P28, Oficinas)
- [ ] Personal cargado (96 personas de la V1)
- [ ] Líderes cargados
- [ ] Maquinaria verificada (MS-001 a MS-008, CH-001 ya cargadas por el schema)
- [ ] Config de actividades configurada (tipo A/B/C)

---

## PASO 6 — Crear usuarios del equipo

Para cada usuario del equipo:
- [ ] Ir a **Authentication → Users → Add user**
  o usar el módulo **Usuarios** dentro de la app
- [ ] Asignar rol correcto en la app

---

## PASO 7 — GitHub

- [ ] Repositorio **privado** `pds-360` creado
- [ ] `.gitignore` incluye `.env.local`
- [ ] Primer commit y push realizados

---

## PASO 8 — Vercel

- [ ] Proyecto importado desde GitHub
- [ ] Variables de entorno configuradas en Vercel:
  - [ ] `NEXT_PUBLIC_SUPABASE_URL`
  - [ ] `NEXT_PUBLIC_SUPABASE_ANON_KEY`
  - [ ] `SUPABASE_SERVICE_ROLE_KEY`
- [ ] Deploy exitoso
- [ ] URL de producción probada
- [ ] Login funciona en producción

---

## PASO 9 — Pruebas con usuarios reales

- [ ] Técnico puede crear planeación con múltiples actividades y líderes distintos
- [ ] Dos técnicos pueden planear simultáneamente sin bloqueos
- [ ] Reporte diario completo (8 pasos) funciona
- [ ] Líder puede aprobar un reporte
- [ ] Dashboard muestra datos correctos
- [ ] Informe exporta a Excel correctamente
- [ ] Cliente solo ve avance (sin nombres de personal)
- [ ] Gerencia ve todo sin poder editar

---

## ✅ Listo para producción cuando todos los pasos estén marcados
