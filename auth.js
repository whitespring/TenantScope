import { broadcastResponseToMainFrame } from './node_modules/@azure/msal-browser/dist/redirect_bridge/index.mjs';

try {
  await broadcastResponseToMainFrame();
} catch {
  document.querySelector('#auth-callback-status').textContent = 'Die Rückgabe konnte nicht verarbeitet werden. Dieses Fenster schließen und die Anmeldung erneut starten.';
}
