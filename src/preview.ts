import * as vscode from 'vscode';

import * as tmlib from './tmlib';

export interface ThreatPreviewData {
  ID: string;
  anchorId: string;
  title: string;
  attack: string;
  impactDesc: string;
  threatType: string;
  fullyMitigated: boolean;
  cvssVector: string;
  isPublic: boolean;
}

export interface ThreatModelPreviewData {
  ID: string;
  anchorId: string;
  title: string;
  analysis: string;
  scopeDescription: string;
  threatCount: number;
  securityObjectiveCount: number;
  attackerCount: number;
}

export type PreviewTarget =
  | {
      kind: 'threat';
      model: ThreatModelPreviewData;
      threat: ThreatPreviewData;
    }
  | {
      kind: 'model';
      model: ThreatModelPreviewData;
      threats: ThreatPreviewData[];
    };

interface RenderedPreview {
  title: string;
  markdown: string;
}

export function toPreviewAnchorId(rawId: string): string {
  const trimmed = (rawId ?? '').trim();
  const base = trimmed.length > 0 ? trimmed : 'item';
  return base.toLowerCase().replace(/[^a-z0-9_-]+/g, '-').replace(/^-+|-+$/g, '');
}

export interface ThreatModelPreviewRenderer {
  readonly id: string;
  readonly label: string;
  render(target: PreviewTarget): RenderedPreview;
}

class ThreatModelPreviewRendererRegistry {
  private readonly renderers = new Map<string, ThreatModelPreviewRenderer>();

  register(renderer: ThreatModelPreviewRenderer): void {
    this.renderers.set(renderer.id, renderer);
  }

  resolve(rendererId?: string): ThreatModelPreviewRenderer {
    if (rendererId) {
      const configuredRenderer = this.renderers.get(rendererId);
      if (configuredRenderer) {
        return configuredRenderer;
      }
    }

    const firstRenderer = this.renderers.values().next().value as ThreatModelPreviewRenderer | undefined;
    if (!firstRenderer) {
      throw new Error('No preview renderers have been registered.');
    }

    return firstRenderer;
  }
}

export class MarkdownThreatModelPreviewRenderer implements ThreatModelPreviewRenderer {
  readonly id = 'markdown-default';
  readonly label = 'Markdown';

  render(target: PreviewTarget): RenderedPreview {
    if (target.kind === 'threat') {
      const title = target.threat.title || target.threat.ID || 'Threat';
      const lines = [
        `<a id="${target.threat.anchorId}"></a>`,
        `# Threat: ${this.escapeInline(title)}`,
        '',
        `- **ID:** ${this.escapeInline(target.threat.ID)}`,
        `- **Type:** ${this.escapeInline(target.threat.threatType)}`,
        `- **CVSS Vector:** ${this.escapeInline(target.threat.cvssVector)}`,
        `- **Fully Mitigated:** ${target.threat.fullyMitigated ? 'Yes' : 'No'}`,
        `- **Public:** ${target.threat.isPublic ? 'Yes' : 'No'}`,
        '',
        '## Attack',
        '',
        this.blockOrFallback(target.threat.attack),
        '',
        '## Impact',
        '',
        this.blockOrFallback(target.threat.impactDesc),
        '',
        '## Model Context',
        '',
        `- **Model:** ${this.escapeInline(target.model.title || target.model.ID)}`,
        `- **Threats In Model:** ${target.model.threatCount}`
      ];

      return {
        title,
        markdown: lines.join('\n')
      };
    }

    const modelTitle = target.model.title || target.model.ID || 'Threat Model';
    const lines = [
      `<a id="${target.model.anchorId}"></a>`,
      `# Threat Model: ${this.escapeInline(modelTitle)}`,
      '',
      `- **ID:** ${this.escapeInline(target.model.ID)}`,
      `- **Threats:** ${target.model.threatCount}`,
      `- **Security Objectives:** ${target.model.securityObjectiveCount}`,
      `- **Attackers:** ${target.model.attackerCount}`,
      '',
      '## Scope',
      '',
      this.blockOrFallback(target.model.scopeDescription),
      '',
      '## Analysis',
      '',
      this.blockOrFallback(target.model.analysis),
      ''
    ];

    if (target.threats.length > 0) {
      lines.push('## Threats', '');
      for (const threat of target.threats) {
        const threatTitle = threat.title || threat.ID || 'Threat';
        lines.push(`<a id="${threat.anchorId}"></a>`);
        lines.push(`### ${this.escapeInline(threatTitle)}`);
        lines.push('');
        lines.push(`- **ID:** ${this.escapeInline(threat.ID)}`);
        lines.push(`- **Type:** ${this.escapeInline(threat.threatType)}`);
        lines.push(`- **Fully Mitigated:** ${threat.fullyMitigated ? 'Yes' : 'No'}`);
        if (threat.attack.trim().length > 0) {
          lines.push(`- **Attack:** ${this.escapeInline(threat.attack)}`);
        }
        if (threat.impactDesc.trim().length > 0) {
          lines.push(`- **Impact:** ${this.escapeInline(threat.impactDesc)}`);
        }
        lines.push('');
      }
    }

    return {
      title: modelTitle,
      markdown: lines.join('\n')
    };
  }

  private blockOrFallback(value: string): string {
    return value.trim().length > 0 ? value : '_No content provided._';
  }

  private escapeInline(value: string): string {
    if (!value || value.length === 0) {
      return '_None_';
    }

    return value.replace(/[\r\n]+/g, ' ').trim();
  }
}

export class ThreatModelPreviewService implements vscode.Disposable, vscode.TextDocumentContentProvider {
  static readonly scheme = 'tmexp-preview';

  private readonly onDidChangeEmitter = new vscode.EventEmitter<vscode.Uri>();
  readonly onDidChange = this.onDidChangeEmitter.event;

  private readonly registry = new ThreatModelPreviewRendererRegistry();
  private readonly documentsByUri = new Map<string, string>();
  private readonly providerRegistration: vscode.Disposable;
  private readonly syncedUri = vscode.Uri.parse(`${ThreatModelPreviewService.scheme}:/synced/active-preview.md`);
  private syncedTargetProvider?: () => PreviewTarget | null;
  private syncedRendererId?: string;

  constructor() {
    this.registry.register(new MarkdownThreatModelPreviewRenderer());
    this.providerRegistration = vscode.workspace.registerTextDocumentContentProvider(ThreatModelPreviewService.scheme, this);
  }

  provideTextDocumentContent(uri: vscode.Uri): string {
    return this.documentsByUri.get(uri.toString()) ?? '# Threat Model Preview\n\n_No preview generated yet._';
  }

  registerRenderer(renderer: ThreatModelPreviewRenderer): void {
    this.registry.register(renderer);
  }

  async showPreview(target: PreviewTarget, rendererId?: string): Promise<void> {
    const renderer = this.registry.resolve(rendererId);
    const rendered = renderer.render(target);

    const fileName = this.toSafeFileName(rendered.title || 'preview');
    const uri = vscode.Uri.parse(`${ThreatModelPreviewService.scheme}:/${renderer.id}/${fileName}.md`);

    this.documentsByUri.set(uri.toString(), rendered.markdown);
    this.onDidChangeEmitter.fire(uri);

    await vscode.commands.executeCommand('markdown.showPreviewToSide', uri);
  }

  async openSyncedPreview(targetProvider: () => PreviewTarget | null, rendererId?: string): Promise<void> {
    this.syncedTargetProvider = targetProvider;
    this.syncedRendererId = rendererId;
    this.refreshSyncedPreview();
    await vscode.commands.executeCommand('markdown.showPreviewToSide', this.syncedUri);
  }

  refreshSyncedPreview(): void {
    const target = this.syncedTargetProvider?.() ?? null;
    if (!target) {
      this.documentsByUri.set(
        this.syncedUri.toString(),
        '# Threat Model Preview\n\n_Select a threat model or threat in the explorer to preview._'
      );
      this.onDidChangeEmitter.fire(this.syncedUri);
      return;
    }

    const renderer = this.registry.resolve(this.syncedRendererId);
    const rendered = renderer.render(target);
    this.documentsByUri.set(this.syncedUri.toString(), rendered.markdown);
    this.onDidChangeEmitter.fire(this.syncedUri);
  }

  isSyncedPreviewActive(): boolean {
    return this.syncedTargetProvider !== undefined;
  }

  async revealSyncedAnchor(anchorId?: string): Promise<void> {
    if (!this.isSyncedPreviewActive()) {
      return;
    }

    const uri = anchorId && anchorId.length > 0
      ? this.syncedUri.with({ fragment: anchorId })
      : this.syncedUri;
    await vscode.commands.executeCommand('markdown.showPreviewToSide', uri);
  }

  createModelPreviewData(model: tmlib.ThreatModel): ThreatModelPreviewData {
    const scopeRecord = (model.scope ?? {}) as unknown as Record<string, unknown>;
    const threats = Array.isArray(model.threats) ? model.threats : [];
    const securityObjectives = Array.isArray(scopeRecord.securityObjectives)
      ? scopeRecord.securityObjectives
      : Array.isArray(model.securityObjectives)
        ? model.securityObjectives
        : [];
    const attackers = Array.isArray(scopeRecord.attackers)
      ? scopeRecord.attackers
      : Array.isArray(model.attackers)
        ? model.attackers
        : [];

    return {
      ID: model.ID ?? '',
      anchorId: toPreviewAnchorId(model.ID ?? model.title ?? 'model'),
      title: model.title ?? '',
      analysis: typeof model.analysis === 'string' ? model.analysis : '',
      scopeDescription: typeof scopeRecord.description === 'string' ? scopeRecord.description : '',
      threatCount: threats.length,
      securityObjectiveCount: securityObjectives.length,
      attackerCount: attackers.length
    };
  }

  createThreatPreviewData(threat: Record<string, unknown>): ThreatPreviewData {
    const cvss = threat.CVSS;
    const cvssVector = cvss && typeof cvss === 'object' && typeof (cvss as Record<string, unknown>).vector === 'string'
      ? (cvss as Record<string, unknown>).vector as string
      : '';

    return {
      ID: this.readString(threat, 'ID'),
      anchorId: toPreviewAnchorId(this.readString(threat, 'ID') || this.readString(threat, 'title') || 'threat'),
      title: this.readString(threat, 'title'),
      attack: this.readString(threat, 'attack'),
      impactDesc: this.readString(threat, 'impactDesc'),
      threatType: this.readString(threat, 'threatType'),
      fullyMitigated: this.readBoolean(threat, 'fullyMitigated'),
      cvssVector,
      isPublic: this.readBoolean(threat, 'public')
    };
  }

  dispose(): void {
    this.providerRegistration.dispose();
    this.onDidChangeEmitter.dispose();
    this.documentsByUri.clear();
    this.syncedTargetProvider = undefined;
    this.syncedRendererId = undefined;
  }

  private readString(input: Record<string, unknown>, key: string): string {
    const value = input[key];
    return typeof value === 'string' ? value : '';
  }

  private readBoolean(input: Record<string, unknown>, key: string): boolean {
    return Boolean(input[key]);
  }

  private toSafeFileName(value: string): string {
    const normalized = value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
    return normalized.length > 0 ? normalized : 'preview';
  }
}
