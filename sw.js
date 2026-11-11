// Service worker minimal : permet l'installation de la page comme une appli.
// Il ne met rien en cache, donc la page est toujours à jour.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
