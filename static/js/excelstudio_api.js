window.ExcelStudioApi = (() => {
  const jsonHeaders = { 'Content-Type': 'application/json; charset=utf-8' };

  const getBaseUrl = () => {
    const el = document.getElementById('apiBaseUrl');
    return (el?.value || '').trim().replace(/\/$/, '');
  };

  async function request(path, options = {}) {
    const url = `${getBaseUrl()}${path}`;
    const response = await fetch(url, options);
    const text = await response.text();
    let data;
    try { data = text ? JSON.parse(text) : {}; } catch { data = { raw: text }; }
    if (!response.ok) {
      const detail = data?.detail || data || { status: response.status };
      throw new Error(typeof detail === 'string' ? detail : JSON.stringify(detail, null, 2));
    }
    return data;
  }

  async function health() {
    return request('/health');
  }

  async function callJson(path, payload) {
    return request(path, { method: 'POST', headers: jsonHeaders, body: JSON.stringify(payload) });
  }

  return {
    getBaseUrl,
    health,
    rulesJson: () => request('/api/rules/discovered'),
    uploadRulesFile: (file) => {
      const form = new FormData();
      form.append('file', file);
      return request('/api/rules/upload', { method: 'POST', body: form });
    },
    clearOutputs: () => request('/api/outputs/clear', { method: 'POST', headers: jsonHeaders }),
    ruleDiscovery: (payload) => callJson('/api/rules/discover', payload),
    ruleDiscoveryJson: (payload) => callJson('/api/rules/discover-json', payload),
    audit: (payload) => callJson('/api/audit', payload),
    auditJson: (payload) => callJson('/api/audit-json', payload),
    markFast: (payload) => callJson('/api/mark-fast', payload),
    rulesOnly: (payload) => callJson('/api/rules-only', payload),
    fullFlow: (payload) => callJson('/api/full-flow', payload),
    fullFlowJson: (payload) => callJson('/api/full-flow-json', payload),
  };
})();
