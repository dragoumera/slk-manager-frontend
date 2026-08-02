const API_URL = import.meta.env.VITE_API_URL || "https://slk-manager-api.onrender.com";

function getToken() {
  return localStorage.getItem("slk_token") || "";
}

async function apiFetch(path, options = {}) {
  const headers = { "Content-Type": "application/json", ...(options.headers || {}) };
  const token = getToken();
  if (token) headers["Authorization"] = "Bearer " + token;
  const res = await fetch(API_URL + path, { ...options, headers });
  if (!res.ok) {
    let msg = "Erreur " + res.status;
    try { const d = await res.json(); if (d && d.error) msg = d.error; } catch (e) {}
    throw new Error(msg);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function login(email, mot_de_passe) {
  const data = await apiFetch("/api/auth/login", {
    method: "POST",
    body: JSON.stringify({ email, mot_de_passe }),
  });
  if (data && data.token) localStorage.setItem("slk_token", data.token);
  return data;
}

function logout() {
  localStorage.removeItem("slk_token");
}

export const api = { apiFetch, login, logout, getToken, API_URL };
