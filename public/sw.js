// Powerchina PDS 360 — Service Worker v2.0
// Offline-first: guarda en cache, sincroniza cuando hay red

const CACHE_NAME   = 'pds360-v3';
const QUEUE_KEY    = 'pds360-offline-queue';
const STATIC_URLS  = ['/', '/icons/icon-192.png', '/icons/icon-512.png', '/manifest.json'];

// ── Instalación ──────────────────────────────────────────────────────────────
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then(c => c.addAll(STATIC_URLS).catch(() => {}))
      .then(() => self.skipWaiting())
  );
});

// ── Activación ───────────────────────────────────────────────────────────────
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// ── Fetch ─────────────────────────────────────────────────────────────────────
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Peticiones a la API: Network-first, si falla → encolar para sync
  if (url.pathname.startsWith('/api/pds360')) {
    e.respondWith(networkFirstWithQueue(e.request));
    return;
  }

  // Documento HTML / navegación: Network-first, para no quedar pegado en una versión vieja tras un deploy
  if (e.request.mode === 'navigate' || e.request.destination === 'document') {
    e.respondWith(
      fetch(e.request).then(res => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => caches.match(e.request).then(cached => cached || new Response('Sin conexión', { status: 503 })))
    );
    return;
  }

  // Resto de recursos estáticos (JS/CSS con hash de build, imágenes): Cache-first
  e.respondWith(
    caches.match(e.request).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (res && res.status === 200 && res.type !== 'opaque') {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => cached || new Response('Sin conexión', { status: 503 }));
    })
  );
});

async function networkFirstWithQueue(request) {
  try {
    const res = await fetch(request.clone());
    return res;
  } catch {
    // Sin red → encolar si es acción de escritura
    try {
      const body = await request.clone().json();
      const writeActions = ['saveSchedule','saveReporteAvance','saveCatalogs','saveMaquinaria','saveBitacora','saveCalendario','saveConfigActividades'];
      if (writeActions.includes(body.action)) {
        await enqueueRequest({ url: request.url, body });
        notifyClients({ type: 'OFFLINE_QUEUED', action: body.action });
        return new Response(JSON.stringify({ ok: true, offline: true, message: 'Guardado localmente. Se sincronizará cuando haya conexión.' }), { headers: { 'Content-Type': 'application/json' } });
      }
    } catch {}
    return new Response(JSON.stringify({ ok: false, error: 'Sin conexión a internet.' }), { status: 503, headers: { 'Content-Type': 'application/json' } });
  }
}

// ── Cola offline ──────────────────────────────────────────────────────────────
async function enqueueRequest(item) {
  const cache = await caches.open(CACHE_NAME);
  const qRes  = await cache.match(QUEUE_KEY).catch(() => null);
  const queue = qRes ? await qRes.json().catch(() => []) : [];
  queue.push({ ...item, ts: Date.now() });
  await cache.put(QUEUE_KEY, new Response(JSON.stringify(queue), { headers: { 'Content-Type': 'application/json' } }));
}

async function flushQueue() {
  const cache = await caches.open(CACHE_NAME);
  const qRes  = await cache.match(QUEUE_KEY).catch(() => null);
  if (!qRes) return;
  const queue = await qRes.json().catch(() => []);
  if (!queue.length) return;

  const remaining = [];
  for (const item of queue) {
    try {
      const res = await fetch(item.url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(item.body) });
      if (!res.ok) remaining.push(item);
    } catch {
      remaining.push(item);
    }
  }

  if (remaining.length) {
    await cache.put(QUEUE_KEY, new Response(JSON.stringify(remaining), { headers: { 'Content-Type': 'application/json' } }));
  } else {
    await cache.delete(QUEUE_KEY);
  }

  notifyClients({ type: 'SYNC_COMPLETE', synced: queue.length - remaining.length, remaining: remaining.length });
}

// ── Sync en background ────────────────────────────────────────────────────────
self.addEventListener('sync', e => {
  if (e.tag === 'pds360-sync') e.waitUntil(flushQueue());
});

// ── Mensaje desde la app ──────────────────────────────────────────────────────
self.addEventListener('message', e => {
  if (e.data?.type === 'FLUSH_QUEUE') e.waitUntil(flushQueue());
  if (e.data?.type === 'GET_QUEUE_SIZE') {
    caches.open(CACHE_NAME).then(async cache => {
      const qRes  = await cache.match(QUEUE_KEY).catch(() => null);
      const queue = qRes ? await qRes.json().catch(() => []) : [];
      e.source.postMessage({ type: 'QUEUE_SIZE', size: queue.length });
    });
  }
});

// ── Online detectado ──────────────────────────────────────────────────────────
self.addEventListener('online', () => flushQueue());

function notifyClients(data) {
  self.clients.matchAll().then(clients => clients.forEach(c => c.postMessage(data)));
}
