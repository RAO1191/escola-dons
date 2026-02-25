// ═══════════════════════════════════════════════════════════════
// ESCOLA DE DONS — Configuração
// ═══════════════════════════════════════════════════════════════

const API_URL = 'https://script.google.com/macros/s/AKfycbxTY2PZPk0TtyzaVwLHpc2epyLpzyywOwqPhOiw-1JABIJZwgUPCiAOpt_zFpsp5Evxwg/exec';

const API = {
  async get(params) {
    const qs = new URLSearchParams(params).toString();
    const res = await fetch(`${API_URL}?${qs}`, { redirect: 'follow' });
    return res.json();
  },
  async post(body) {
    const res = await fetch(API_URL, {
      method: 'POST',
      redirect: 'follow',
      body: JSON.stringify(body)
    });
    return res.json();
  }
};

async function loadSiteImages() {
  try {
    const res = await API.get({ action: 'getConfig' });
    if (!res.ok) return;
    if (res.bgFileId) {
      const url = `https://drive.google.com/uc?export=view&id=${res.bgFileId}`;
      document.querySelectorAll('.site-bg').forEach(el => {
        el.style.backgroundImage = `url('${url}')`;
        el.style.backgroundSize = 'cover';
        el.style.backgroundPosition = 'center';
      });
    }
    if (res.logoFileId) {
      const url = `https://drive.google.com/uc?export=view&id=${res.logoFileId}`;
      document.querySelectorAll('.site-logo').forEach(el => { el.src = url; el.style.display = 'block'; });
      document.querySelectorAll('.site-logo-placeholder').forEach(el => el.style.display = 'none');
    }
    if (res.churchName) {
      document.querySelectorAll('.church-name').forEach(el => el.textContent = res.churchName);
    }
  } catch(e) {}
}
