// Communication avec l'API SLK Manager (backend sur Render).
// L'adresse de l'API est lue depuis une variable d'environnement définie au
// moment du build (VITE_API_URL). Si elle est absente, on retombe sur
// l'adresse Render par défaut — pratique pour un test rapide.

const API_URL = (import.meta.env.VITE_API_URL || "https://slk-manager-api.onrender.com").replace(/\/+$/, "");

// Récupère le token stocké après connexion (en mémoire de session du navigateur).
function getToken() {
  try {
    return window.sessionStorage.getItem("slk_token");
  } catch {
    return null;
  }
}

function setToken(token) {
  try {
    if (token) window.sessionStorage.setItem("slk_token", token);
    else window.sessionStorage.removeItem("slk_token");
  } catch {
    /* sessionStorage indisponible : on ignore, la session durera le temps de l'onglet */
  }
}

// Appel générique à l'API. Ajoute automatiquement le token d'authentification
// si présent, et renvoie le JSON (ou lève une erreur lisible).
async function apiFetch(chemin, options = {}) {
  const headers = { "Content-Type": "application/json", ...(options.headers || {}) };
  const token = getToken();
  if (token) headers.Authorization = "Bearer " + token;

  let reponse;
  try {
    reponse = await fetch(API_URL + chemin, { ...options, headers });
  } catch (e) {
    throw new Error(
      "Impossible de joindre le serveur. Si c'est le premier appel depuis un moment, " +
      "il peut mettre jusqu'à 50 secondes à se réveiller — réessayez dans un instant."
    );
  }

  let data = null;
  const texte = await reponse.text();
  if (texte) {
    try { data = JSON.parse(texte); } catch { data = { brut: texte }; }
  }

  if (!reponse.ok) {
    const message = (data && data.error) || "Erreur " + reponse.status;
    throw new Error(message);
  }
  return data;
}

// Connexion : renvoie l'utilisateur et mémorise le token.
async function login(email, mot_de_passe) {
  const data = await apiFetch("/api/auth/login", {
    method: "POST",
    body: JSON.stringify({ email, mot_de_passe }),
  });
  if (data && data.token) setToken(data.token);
  return data; // { token, user }
}

function logout() {
  setToken(null);
}

export const api = { apiFetch, login, logout, getToken, API_URL };
