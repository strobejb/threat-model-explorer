import {
  AttackerEditorPayload,
  EditorEntityKind,
  EntityEditorPayload,
  EntityEditorViewModel,
  SecurityObjectiveEditorPayload,
  ThreatEditorPayload,
  ThreatModelEditorPayload
} from './types';

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

const commitSvg = '<svg viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="M13.78 4.22a.75.75 0 0 1 0 1.06l-7.25 7.25a.75.75 0 0 1-1.06 0L2.22 9.28a.75.75 0 1 1 1.06-1.06L6 10.94l6.72-6.72a.75.75 0 0 1 1.06 0z" fill="currentColor"/></svg>';

function textInput(id: string, value: string): string {
  return `<div class="text-field-wrapper"><input id="${id}" type="text" value="${value}" /><button type="button" class="commit-btn" data-field="${id}" title="Commit change" aria-label="Commit change">${commitSvg}</button></div>`;
}

function getSharedStyles(): string {
  return `
    body { font-family: var(--vscode-font-family); padding: 16px; color: var(--vscode-foreground); }
    .panel-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
    .panel-header h2 { margin: 0; }
    .grid { display: grid; gap: 12px; }
    label { display: block; font-weight: 600; margin-bottom: 4px; }
    input[type="text"], textarea {
      width: 100%; box-sizing: border-box;
      background: var(--vscode-input-background);
      color: var(--vscode-input-foreground);
      border: 1px solid var(--vscode-input-border);
      padding: 8px;
    }
    textarea { min-height: 84px; resize: vertical; }
    .input-with-action { display: grid; grid-template-columns: 1fr auto; gap: 8px; }
    .icon-button {
      width: 36px;
      height: 36px;
      padding: 0;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      border: 1px solid var(--vscode-button-border, transparent);
      background: var(--vscode-button-secondaryBackground);
      color: var(--vscode-button-secondaryForeground);
      transition: background-color 120ms ease, transform 120ms ease, box-shadow 120ms ease, border-color 120ms ease;
    }
    .icon-button:hover {
      background: var(--vscode-button-secondaryHoverBackground);
      border-color: var(--vscode-focusBorder);
      transform: translateY(-1px);
      box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2);
    }
    .icon-button:focus-visible {
      outline: 1px solid var(--vscode-focusBorder);
      outline-offset: 1px;
    }
    .icon-button:active {
      transform: translateY(0);
      box-shadow: none;
    }
    .icon-button svg { width: 18px; height: 18px; }
    .text-field-wrapper { position: relative; }
    .text-field-wrapper input[type="text"] { padding-right: 30px; }
    .commit-btn {
      position: absolute;
      right: 4px;
      top: 50%;
      transform: translateY(-50%);
      display: none;
      align-items: center;
      justify-content: center;
      width: 22px;
      height: 22px;
      padding: 0;
      border: none;
      border-radius: 3px;
      background: transparent;
      color: var(--vscode-input-foreground);
      cursor: pointer;
      opacity: 0.7;
    }
    .commit-btn:hover { opacity: 1; background: var(--vscode-toolbar-hoverBackground); }
    .commit-btn.visible { display: inline-flex; }
    .commit-btn svg { width: 14px; height: 14px; }
    #ID { font-family: var(--vscode-editor-font-family); }
    #attack { font-family: var(--vscode-font-family); }
    #impactDesc { font-family: var(--vscode-font-family); }
    .checkbox { display: flex; align-items: center; gap: 8px; }
    button { padding: 8px 14px; }
    dialog { border: 1px solid var(--vscode-panel-border); background: var(--vscode-editor-background); color: var(--vscode-foreground); padding: 0; width: min(760px, 90vw); }
    dialog::backdrop { background: rgba(0, 0, 0, 0.45); }
    .modal-body { padding: 14px; }
    .metric-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
    .metric { border: none; padding: 0; border-radius: 0; }
    .metric h4 { margin: 0 0 6px 0; font-size: 12px; }
    .metric select { width: 100%; background: var(--vscode-dropdown-background); color: var(--vscode-dropdown-foreground); border: none; outline: none; padding: 6px; box-shadow: none; }
    .cvss-output { margin-top: 10px; padding: 8px; border: 1px solid var(--vscode-panel-border); }
    .modal-actions { margin-top: 12px; display: flex; gap: 8px; justify-content: flex-end; }
  `;
}

function getThreatBody(payload: ThreatEditorPayload): string {
  const p = {
    ID: escapeHtml(payload.ID),
    title: escapeHtml(payload.title),
    attack: escapeHtml(payload.attack),
    impactDesc: escapeHtml(payload.impactDesc),
    threatType: escapeHtml(payload.threatType),
    cvssVector: escapeHtml(payload.cvssVector),
  };

  return `
  <div class="grid">
    <div><label for="ID">ID</label>${textInput('ID', p.ID)}</div>
    <div><label for="title">Title</label>${textInput('title', p.title)}</div>
    <div><label for="attack">Attack</label><textarea id="attack">${p.attack}</textarea></div>
    <div><label for="impactDesc">Impact Description</label><textarea id="impactDesc">${p.impactDesc}</textarea></div>
    <div><label for="threatType">Threat Type</label>${textInput('threatType', p.threatType)}</div>
    <div>
      <label for="cvssVector">CVSS Vector</label>
      <div class="input-with-action">
        ${textInput('cvssVector', p.cvssVector)}
        <button type="button" id="cvssCalcOpen" class="icon-button" title="Open CVSS Calculator" aria-label="Open CVSS Calculator">
          <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
            <rect x="5" y="2.5" width="14" height="19" rx="2" ry="2" fill="none" stroke="currentColor" stroke-width="1.6"/>
            <rect x="8" y="5.5" width="8" height="3" rx="0.8" fill="currentColor" opacity="0.8"/>
            <circle cx="9" cy="12" r="1" fill="currentColor"/>
            <circle cx="12" cy="12" r="1" fill="currentColor"/>
            <circle cx="15" cy="12" r="1" fill="currentColor"/>
            <circle cx="9" cy="15.5" r="1" fill="currentColor"/>
            <circle cx="12" cy="15.5" r="1" fill="currentColor"/>
            <circle cx="15" cy="15.5" r="1" fill="currentColor"/>
            <circle cx="9" cy="19" r="1" fill="currentColor"/>
            <rect x="11" y="18" width="5" height="2" rx="0.7" fill="currentColor"/>
          </svg>
        </button>
      </div>
    </div>
    <label class="checkbox"><input id="fullyMitigated" type="checkbox" ${payload.fullyMitigated ? 'checked' : ''}/> Fully Mitigated</label>
    <label class="checkbox"><input id="public" type="checkbox" ${payload.public ? 'checked' : ''}/> Public</label>
  </div>

  <dialog id="cvssModal">
    <div class="modal-body">
      <h3>CVSS v3.1 Base Calculator</h3>
      <div class="metric-grid">
        <div class="metric"><h4>Attack Vector (AV)</h4><select id="mAV"></select></div>
        <div class="metric"><h4>Attack Complexity (AC)</h4><select id="mAC"></select></div>
        <div class="metric"><h4>Privileges Required (PR)</h4><select id="mPR"></select></div>
        <div class="metric"><h4>User Interaction (UI)</h4><select id="mUI"></select></div>
        <div class="metric"><h4>Scope (S)</h4><select id="mS"></select></div>
        <div class="metric"><h4>Confidentiality (C)</h4><select id="mC"></select></div>
        <div class="metric"><h4>Integrity (I)</h4><select id="mI"></select></div>
        <div class="metric"><h4>Availability (A)</h4><select id="mA"></select></div>
      </div>
      <div class="cvss-output">
        <div><strong>Vector:</strong> <span id="cvssVectorPreview"></span></div>
        <div><strong>Base Score:</strong> <span id="cvssScorePreview"></span></div>
      </div>
      <div class="modal-actions">
        <button type="button" id="cvssCancel">Cancel</button>
        <button type="button" id="cvssApply">Apply</button>
      </div>
    </div>
  </dialog>
`;
}

function getSecurityObjectiveBody(payload: SecurityObjectiveEditorPayload): string {
  return `
  <div class="grid">
    <div><label for="ID">ID</label>${textInput('ID', escapeHtml(payload.ID))}</div>
    <div><label for="title">Title</label>${textInput('title', escapeHtml(payload.title))}</div>
    <div><label for="description">Description</label><textarea id="description">${escapeHtml(payload.description)}</textarea></div>
  </div>
`;
}

function getAttackerBody(payload: AttackerEditorPayload): string {
  return `
  <div class="grid">
    <div><label for="ID">ID</label>${textInput('ID', escapeHtml(payload.ID))}</div>
    <div><label for="name">Name</label>${textInput('name', escapeHtml(payload.name))}</div>
    <div><label for="title">Title</label>${textInput('title', escapeHtml(payload.title))}</div>
    <div><label for="description">Description</label><textarea id="description">${escapeHtml(payload.description)}</textarea></div>
  </div>
`;
}

function getThreatModelBody(payload: ThreatModelEditorPayload): string {
  return `
  <div class="grid">
    <div><label for="ID">ID</label>${textInput('ID', escapeHtml(payload.ID))}</div>
    <div><label for="title">Title</label>${textInput('title', escapeHtml(payload.title))}</div>
    <div><label for="version">Version</label>${textInput('version', escapeHtml(payload.version))}</div>
    <div><label for="analysis">Analysis</label><textarea id="analysis">${escapeHtml(payload.analysis)}</textarea></div>
  </div>
`;
}

function getBodyByKind(kind: EditorEntityKind, payload: EntityEditorPayload): string {
  if (kind === 'threat') {
    return getThreatBody(payload as ThreatEditorPayload);
  }

  if (kind === 'securityObjective') {
    return getSecurityObjectiveBody(payload as SecurityObjectiveEditorPayload);
  }

  if (kind === 'model') {
    return getThreatModelBody(payload as ThreatModelEditorPayload);
  }

  return getAttackerBody(payload as AttackerEditorPayload);
}

function getFieldIds(kind: EditorEntityKind): { all: string[]; enter: string[]; checkboxes: string[] } {
  if (kind === 'threat') {
    return {
      all: ['ID', 'title', 'attack', 'impactDesc', 'threatType', 'cvssVector', 'fullyMitigated', 'public'],
      enter: ['ID', 'title', 'threatType', 'cvssVector'],
      checkboxes: ['fullyMitigated', 'public']
    };
  }

  if (kind === 'securityObjective') {
    return {
      all: ['ID', 'title', 'description'],
      enter: ['ID', 'title'],
      checkboxes: []
    };
  }

  if (kind === 'model') {
    return {
      all: ['ID', 'title', 'version', 'analysis'],
      enter: ['ID', 'title', 'version'],
      checkboxes: []
    };
  }

  return {
    all: ['ID', 'name', 'title', 'description'],
    enter: ['ID', 'name', 'title'],
    checkboxes: []
  };
}

function getSavePayloadScript(kind: EditorEntityKind): string {
  if (kind === 'threat') {
    return `{
      ID: document.getElementById('ID').value,
      title: document.getElementById('title').value,
      attack: document.getElementById('attack').value,
      impactDesc: document.getElementById('impactDesc').value,
      threatType: document.getElementById('threatType').value,
      cvssVector: document.getElementById('cvssVector').value,
      fullyMitigated: document.getElementById('fullyMitigated').checked,
      public: document.getElementById('public').checked
    }`;
  }

  if (kind === 'securityObjective') {
    return `{
      ID: document.getElementById('ID').value,
      title: document.getElementById('title').value,
      description: document.getElementById('description').value
    }`;
  }

  if (kind === 'model') {
    return `{
      ID: document.getElementById('ID').value,
      title: document.getElementById('title').value,
      version: document.getElementById('version').value,
      analysis: document.getElementById('analysis').value
    }`;
  }

  return `{
    ID: document.getElementById('ID').value,
    name: document.getElementById('name').value,
    title: document.getElementById('title').value,
    description: document.getElementById('description').value
  }`;
}

function getCvssScript(kind: EditorEntityKind): string {
  if (kind !== 'threat') {
    return '';
  }

  return `
    const cvssVectorInput = document.getElementById('cvssVector');
    const cvssModal = document.getElementById('cvssModal');

    const metricDefs = {
      AV: [
        { code: 'N', label: 'Network', value: 0.85 },
        { code: 'A', label: 'Adjacent', value: 0.62 },
        { code: 'L', label: 'Local', value: 0.55 },
        { code: 'P', label: 'Physical', value: 0.20 }
      ],
      AC: [
        { code: 'L', label: 'Low', value: 0.77 },
        { code: 'H', label: 'High', value: 0.44 }
      ],
      PR: [
        { code: 'N', label: 'None', valueU: 0.85, valueC: 0.85 },
        { code: 'L', label: 'Low', valueU: 0.62, valueC: 0.68 },
        { code: 'H', label: 'High', valueU: 0.27, valueC: 0.50 }
      ],
      UI: [
        { code: 'N', label: 'None', value: 0.85 },
        { code: 'R', label: 'Required', value: 0.62 }
      ],
      S: [
        { code: 'U', label: 'Unchanged' },
        { code: 'C', label: 'Changed' }
      ],
      C: [
        { code: 'N', label: 'None', value: 0.00 },
        { code: 'L', label: 'Low', value: 0.22 },
        { code: 'H', label: 'High', value: 0.56 }
      ],
      I: [
        { code: 'N', label: 'None', value: 0.00 },
        { code: 'L', label: 'Low', value: 0.22 },
        { code: 'H', label: 'High', value: 0.56 }
      ],
      A: [
        { code: 'N', label: 'None', value: 0.00 },
        { code: 'L', label: 'Low', value: 0.22 },
        { code: 'H', label: 'High', value: 0.56 }
      ]
    };

    function fillMetricSelect(metricId, options) {
      const select = document.getElementById('m' + metricId);
      options.forEach((option) => {
        const el = document.createElement('option');
        el.value = option.code;
        el.textContent = option.code + ' - ' + option.label;
        select.appendChild(el);
      });
    }

    Object.keys(metricDefs).forEach((metricId) => fillMetricSelect(metricId, metricDefs[metricId]));

    function getMetric(metricId, code) {
      return metricDefs[metricId].find((m) => m.code === code);
    }

    function roundUp1(value) {
      return Math.ceil(value * 10) / 10;
    }

    function computeCvss() {
      const AV = document.getElementById('mAV').value;
      const AC = document.getElementById('mAC').value;
      const PR = document.getElementById('mPR').value;
      const UI = document.getElementById('mUI').value;
      const S = document.getElementById('mS').value;
      const C = document.getElementById('mC').value;
      const I = document.getElementById('mI').value;
      const A = document.getElementById('mA').value;

      const av = getMetric('AV', AV).value;
      const ac = getMetric('AC', AC).value;
      const prMetric = getMetric('PR', PR);
      const pr = S === 'C' ? prMetric.valueC : prMetric.valueU;
      const ui = getMetric('UI', UI).value;
      const c = getMetric('C', C).value;
      const i = getMetric('I', I).value;
      const a = getMetric('A', A).value;

      const iss = 1 - ((1 - c) * (1 - i) * (1 - a));
      const impact = S === 'U'
        ? 6.42 * iss
        : 7.52 * (iss - 0.029) - 3.25 * Math.pow(iss - 0.02, 15);
      const exploitability = 8.22 * av * ac * pr * ui;

      let score = 0;
      if (impact > 0) {
        const base = S === 'U' ? impact + exploitability : 1.08 * (impact + exploitability);
        score = Math.min(base, 10);
      }
      score = roundUp1(score);

      const vector = 'CVSS:3.1/AV:' + AV + '/AC:' + AC + '/PR:' + PR + '/UI:' + UI + '/S:' + S + '/C:' + C + '/I:' + I + '/A:' + A;
      document.getElementById('cvssVectorPreview').textContent = vector;
      document.getElementById('cvssScorePreview').textContent = score.toFixed(1);
      return { vector, score };
    }

    function parseVector(vector) {
      const defaults = { AV: 'N', AC: 'L', PR: 'N', UI: 'N', S: 'U', C: 'N', I: 'N', A: 'N' };
      if (!vector || typeof vector !== 'string') {
        return defaults;
      }
      const parts = vector.split('/');
      parts.forEach((part) => {
        const [k, v] = part.split(':');
        if (defaults[k] !== undefined && typeof v === 'string') {
          defaults[k] = v;
        }
      });
      return defaults;
    }

    function applyParsedMetrics(metrics) {
      Object.keys(metrics).forEach((metricId) => {
        const select = document.getElementById('m' + metricId);
        if (select) {
          select.value = metrics[metricId];
        }
      });
      computeCvss();
    }

    ['mAV', 'mAC', 'mPR', 'mUI', 'mS', 'mC', 'mI', 'mA'].forEach((id) => {
      document.getElementById(id).addEventListener('change', computeCvss);
    });

    document.getElementById('cvssCalcOpen').addEventListener('click', () => {
      applyParsedMetrics(parseVector(cvssVectorInput.value));
      cvssModal.showModal();
    });

    document.getElementById('cvssCancel').addEventListener('click', () => {
      cvssModal.close();
    });

    document.getElementById('cvssApply').addEventListener('click', () => {
      const calculated = computeCvss();
      cvssVectorInput.value = calculated.vector;
      updateCommitButton('cvssVector');
      cvssModal.close();
    });
  `;
}

export function renderEntityEditorHtml(model: EntityEditorViewModel): string {
  const initialFocus = escapeHtml(model.focusField ?? '');
  const fieldIds = getFieldIds(model.kind);
  const fieldIdsJson = JSON.stringify(fieldIds.all);
  const enterFieldIdsJson = JSON.stringify(fieldIds.enter);
  const checkboxFieldIdsJson = JSON.stringify(fieldIds.checkboxes);
  const savePayloadScript = getSavePayloadScript(model.kind);

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Edit ${model.kind}</title>
  <style>
${getSharedStyles()}
  </style>
</head>
<body>
  <div class="panel-header">
    <h2>Edit ${model.kind === 'securityObjective' ? 'Security Objective' : model.kind === 'attacker' ? 'Attacker' : model.kind === 'model' ? 'Threat Model' : 'Threat'}</h2>
    <button id="save">Update</button>
  </div>
${getBodyByKind(model.kind, model.payload)}

  <script>
    const vscode = acquireVsCodeApi();
    const initialFocusField = '${initialFocus}';
    const fieldIds = ${fieldIdsJson};
    const enterFieldIds = ${enterFieldIdsJson};
    const checkboxFieldIds = ${checkboxFieldIdsJson};

${getCvssScript(model.kind)}

    fieldIds.forEach((fieldId) => {
      const el = document.getElementById(fieldId);
      if (!el) {
        return;
      }

      const reveal = () => {
        vscode.postMessage({ type: 'revealField', fieldId });
      };

      el.addEventListener('focus', reveal);
      el.addEventListener('click', reveal);
    });

    checkboxFieldIds.forEach((fieldId) => {
      const el = document.getElementById(fieldId);
      if (!el) {
        return;
      }

      el.addEventListener('change', () => {
        vscode.postMessage({ type: 'updateField', fieldId, value: el.checked });
      });
    });

    const originals = {};
    enterFieldIds.forEach((fieldId) => {
      const el = document.getElementById(fieldId);
      if (el) { originals[fieldId] = el.value; }
    });

    function updateCommitButton(fieldId) {
      const el = document.getElementById(fieldId);
      const btn = document.querySelector('.commit-btn[data-field="' + fieldId + '"');
      if (!el || !btn) { return; }
      if (el.value !== originals[fieldId]) {
        btn.classList.add('visible');
      } else {
        btn.classList.remove('visible');
      }
    }

    function commitField(fieldId) {
      const el = document.getElementById(fieldId);
      if (!el || el.value === originals[fieldId]) { return; }
      vscode.postMessage({ type: 'updateField', fieldId, value: el.value });
      originals[fieldId] = el.value;
      updateCommitButton(fieldId);
    }

    function resetDirtyState() {
      enterFieldIds.forEach((fieldId) => {
        const el = document.getElementById(fieldId);
        if (el) {
          originals[fieldId] = el.value;
          updateCommitButton(fieldId);
        }
      });
    }

    function submitUpdate() {
      vscode.postMessage({
        type: 'save',
        payload: ${savePayloadScript}
      });
      resetDirtyState();
    }

    document.getElementById('save').addEventListener('click', () => {
      submitUpdate();
    });

    enterFieldIds.forEach((fieldId) => {
      const el = document.getElementById(fieldId);
      if (!el) { return; }

      el.addEventListener('input', () => updateCommitButton(fieldId));

      el.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
          event.preventDefault();
          commitField(fieldId);
        }
      });
    });

    document.querySelectorAll('.commit-btn').forEach((btn) => {
      btn.addEventListener('mousedown', (e) => e.preventDefault());
      btn.addEventListener('click', () => commitField(btn.dataset.field));
    });

    if (initialFocusField) {
      const initialEl = document.getElementById(initialFocusField);
      if (initialEl) {
        setTimeout(() => {
          initialEl.focus();
          if (typeof initialEl.select === 'function') {
            initialEl.select();
          }
        }, 0);
      }
    }
  </script>
</body>
</html>`;
}

export function renderEmptyEditorHtml(): string {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Threat Editor</title>
  <style>
    body { font-family: var(--vscode-font-family); padding: 16px; color: var(--vscode-foreground); }
  </style>
</head>
<body>
  <h2>Threat Editor</h2>
  <p>Select a threat, security objective, or attacker in the explorer to edit fields.</p>
</body>
</html>`;
}
