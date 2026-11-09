const CACHE_NAME = 'qr-event-demo-v6';
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './styles.css?v=6.0',
    './app.js?v=6.0',
    './manifest.json',
    'https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,500;0,600;0,700;1,400;1,600&family=Urbanist:wght@400;500;600;700;800&display=swap',
    'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/tabler-icons.min.css',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.min.js',
    'https://cdn.jsdelivr.net/npm/qrcode@1.5.3/build/qrcode.min.js',
    'https://cdn.jsdelivr.net/npm/canvas-confetti@1.6.0/dist/confetti.browser.min.js'
];

self.addEventListener('install', (e) => {
    e.waitUntil(
        caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS_TO_CACHE).catch(err => console.warn('Cache error:', err)))
    );
    self.skipWaiting();
});

self.addEventListener('activate', (e) => {
    e.waitUntil(
        caches.keys().then((keys) => {
            return Promise.all(
                keys.map((key) => {
                    if (key !== CACHE_NAME) {
                        console.log('[SW Demo] Deleting old cache:', key);
                        return caches.delete(key);
                    }
                })
            );
        })
    );
    self.clients.claim();
});

// Network-First para archivos locales
self.addEventListener('fetch', (e) => {
    if (!e.request.url.startsWith('http')) return;

    const url = new URL(e.request.url);
    const isLocal = url.origin === self.location.origin;

    if (isLocal) {
        e.respondWith(
            fetch(e.request)
                .then(networkRes => {
                    if (networkRes && networkRes.status === 200) {
                        const resClone = networkRes.clone();
                        caches.open(CACHE_NAME).then(c => c.put(e.request, resClone));
                    }
                    return networkRes;
                })
                .catch(() => caches.match(e.request).then(cached => cached || caches.match('./index.html')))
        );
        return;
    }

    // Librerías externas
    e.respondWith(
        caches.match(e.request).then((res) => res || fetch(e.request))
    );
});
