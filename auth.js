// Accesso a Microsoft (OneDrive) tramite MSAL, flusso "redirect" (adatto anche a iPhone).
import { CONFIG } from "./config.js";

const SCOPES = ["Files.ReadWrite"];
let pca = null;

export function getClientId() {
  try { return CONFIG.clientId || localStorage.getItem("gp_clientId") || ""; } catch { return CONFIG.clientId || ""; }
}
export function saveClientId(id) {
  try { localStorage.setItem("gp_clientId", id.trim()); } catch {}
}

export async function initAuth() {
  const clientId = getClientId();
  if (!clientId) return { state: "noclient" };
  pca = new msal.PublicClientApplication({
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: location.origin + location.pathname.replace(/[^/]*$/, ""),
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: "localStorage" },
  });
  await pca.initialize();
  const res = await pca.handleRedirectPromise();
  if (res && res.account) pca.setActiveAccount(res.account);
  if (!pca.getActiveAccount()) {
    const all = pca.getAllAccounts();
    if (all.length) pca.setActiveAccount(all[0]);
  }
  const acc = pca.getActiveAccount();
  return acc ? { state: "ok", account: acc } : { state: "loggedout" };
}

export function login() {
  return pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
}

export async function logout() {
  const acc = pca.getActiveAccount();
  await pca.logoutRedirect({ account: acc, postLogoutRedirectUri: location.origin + location.pathname });
}

export async function getToken() {
  const account = pca.getActiveAccount();
  try {
    const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    await pca.acquireTokenRedirect({ scopes: SCOPES, account });
    throw new Error("Rinnovo accesso in corso…");
  }
}
