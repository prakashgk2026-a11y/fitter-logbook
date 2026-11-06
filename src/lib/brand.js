// The name shown in the menu and browser tab. Set VITE_APP_NAME (GitHub → Variables, or .env.production)
// to rename the whole app, e.g. "Fitter Logbook". Defaults to "ITI Logbook".
export const APP_NAME = (import.meta.env.VITE_APP_NAME || '').trim() || 'ITI Logbook'
