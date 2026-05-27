# Powerchina · PDS 360 — v2.0 (Supabase Edition)

**Gestión integral de obra · PostgreSQL real · Sin bloqueos · Sin timeouts**

---

## Por qué Supabase vs Apps Script

| Problema V1 | Solución V2 |
|---|---|
| Bloqueos con múltiples usuarios simultáneos | PostgreSQL maneja miles de peticiones simultáneas |
| Timeout al llenar formularios largos | Sesión con auto-refresh, nunca expira |
| Un solo líder por especialidad | Cada actividad tiene su propio líder (columna separada) |
| Consultas lentas con muchos datos | Índices reales en PostgreSQL |

---

## Stack

- **Frontend:** Next.js 14 · TypeScript · Tailwind CSS
- **Base de datos:** Supabase (PostgreSQL)
- **Auth:** Supabase Auth (JWT real, sesiones duraderas)
- **Deploy:** Vercel
- **Offline:** Service Worker
- **Costo:** $0

---

## Instalación paso a paso

### 1. Crear proyecto Next.js

```powershell
cd C:\Users\USUARIO\Documents\Proyectos
npx create-next-app@latest pds-360 --typescript --tailwind --app --no-src-dir --no-import-alias
cd pds-360
npm install @supabase/supabase-js xlsx
```

### 2. Copiar los archivos de este ZIP

Copia todos los archivos al proyecto `pds-360` manteniendo la estructura de carpetas.

### 3. Configurar variables de entorno

```powershell
copy .env.local.example .env.local
```

El `.env.local.example` ya tiene las claves correctas. Solo cópialo como `.env.local`.

### 4. Correr en desarrollo

```powershell
npm run dev
```

Abre http://localhost:3000

---

## Configurar Supabase

### Crear las tablas

1. Ve a **supabase.com** → tu proyecto `pds-360`
2. Menú izquierdo → **SQL Editor**
3. Clic en **New query**
4. Copia y pega todo el contenido de `supabase/schema.sql`
5. Clic en **Run** (o `Ctrl+Enter`)
6. Verifica que dice **"Success"** al final

### Crear el primer usuario admin

1. En Supabase → **Authentication → Users → Add user**
2. Email: `dcetinag@gmail.com`
3. Password: la que quieras
4. Clic en **Create user**
5. Ve a **SQL Editor** y ejecuta:

```sql
update profiles
set rol = 'admin', nombre = 'David Cetina'
where correo = 'dcetinag@gmail.com';
```

### Verificar las tablas

En Supabase → **Table Editor** → deberías ver 20+ tablas incluyendo:
- `profiles`
- `especialidades_actividades`
- `areas`
- `personal`
- `programaciones`
- `reportes_avance`
- etc.

---

## Roles del sistema

| Rol | Puede hacer |
|---|---|
| **admin** | Todo: catálogos, usuarios, planeación, reportes, aprobación |
| **lider** | Planea su especialidad, aprueba reportes de sus técnicos |
| **tecnico** | Su propia planeación y reporte diario |
| **gerencia** | Ve todo, no edita nada |
| **cliente** | Solo avance del proyecto |
| **visualizador** | Reportes e informes, sin editar |

---

## Cargar datos iniciales (catálogos de la V1)

1. Inicia sesión con el usuario admin
2. Ve a **Catálogos**
3. Para cada catálogo (Especialidades, Áreas, Personal, Líderes):
   - Clic en **Cargar Excel**
   - Sube el archivo correspondiente
   - Confirma el reemplazo

También puedes cargar directamente en Supabase:
- **Table Editor** → selecciona la tabla → **Insert rows**

---

## Deploy en Vercel

```powershell
git init
git add .
git commit -m "PDS 360 v2.0 - Supabase"
git remote add origin https://github.com/dcetinag/pds-360.git
git push -u origin main
```

En vercel.com:
1. **Add New Project** → importa `pds-360`
2. **Environment Variables** → agrega las 3 variables de `.env.local`
3. **Deploy**

---

## Seguridad

```
· Repositorio GitHub: privado, solo tu cuenta
· Supabase: proyecto en dcetinag@gmail.com
· Vercel: cuenta personal
· Service Role Key: solo en variables de Vercel (servidor)
  nunca en el código del browser

Para revocar acceso a Powerchina:
→ Desactiva usuarios desde el módulo Usuarios
→ O elimina el proyecto en Supabase
→ La app deja de funcionar inmediatamente
```

---

## Propiedad

Desarrollado por **David Cetina** (dcetinag@gmail.com).
Proyecto, código y datos son propiedad exclusiva de David Cetina.
