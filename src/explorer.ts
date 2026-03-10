import * as vscode from 'vscode';
import * as path from 'path';
import * as tmlib from './tmlib';
import * as YAML from 'yaml';
import fs from 'fs';
import { renderEmptyEditorHtml, renderEntityEditorHtml } from './panel/editorHtml';
import {
  AttackerEditorPayload,
  EditorEntityKind,
  EntityEditorPayload,
  SecurityObjectiveEditorPayload,
  ThreatEditorPayload,
  ThreatModelEditorPayload
} from './panel/types';

type TreeItemCategory = 'threat' | 'securityObjective' | 'attacker' | 'model' | 'other';

export interface ParsedModel {
  source: string;
  doc: YAML.Document;
  model: tmlib.ThreatModel;
}

export interface ModelNode {
  id: string;
  title: string;
  filePath: string;
  doc: YAML.Document;
  model: tmlib.ThreatModel;
  source: string;
  children: ModelNode[];
}

export
class YAMLTreeDataProvider implements vscode.TreeDataProvider<YAMLTreeItem>, vscode.WebviewViewProvider {
  private _onDidChangeTreeData: vscode.EventEmitter<YAMLTreeItem | undefined | null | void> = new vscode.EventEmitter<YAMLTreeItem | undefined | null | void>();
  readonly onDidChangeTreeData: vscode.Event<YAMLTreeItem | undefined | null | void> = this._onDidChangeTreeData.event;

  private model?: tmlib.ThreatModel;
  private tmDoc?: YAML.Document;
  private tmSource?: string;
  private editorView?: vscode.WebviewView;
  private activeEntityKind?: EditorEntityKind;
  private activeEntityMap?: YAML.YAMLMap;
  private activeEntityDraft?: EntityEditorPayload;
  private activeEntityIsDraft = false;
  private activeEntityLabel = 'Threat';
  private activeEntityIndex?: number;
  private activeEntityId?: string;
  private suppressNextExternalChange = false;
  private pendingFocusField?: string;
  private modelTree: ModelNode | null = null;
  private modelsByPath: Map<string, ModelNode> = new Map();
  private treeView?: vscode.TreeView<YAMLTreeItem>;
  private treeItemsByModelPath: Map<string, YAMLTreeItem> = new Map();
  private pendingRevealPath?: string;
  private parsedCache: Map<string, ParsedModel> = new Map();
  private suppressRevealSelection = false;
  private editorSwitchSuppressCount = 0;
  constructor(private workspaceRoot: string = '') {
  }

  resolveWebviewView(webviewView: vscode.WebviewView): void {
    this.editorView = webviewView;
    webviewView.webview.options = { enableScripts: true };
    if (this.activeEntityKind && this.activeEntityMap) {
      webviewView.webview.html = this.getEntityEditorHtml(this.activeEntityKind, this.activeEntityMap, this.activeEntityLabel, this.pendingFocusField);
      this.pendingFocusField = undefined;
    } else if (this.activeEntityKind && this.activeEntityDraft) {
      webviewView.webview.html = this.getEntityEditorHtmlFromPayload(this.activeEntityKind, this.activeEntityDraft, this.activeEntityLabel, this.pendingFocusField);
      this.pendingFocusField = undefined;
    } else {
      webviewView.webview.html = renderEmptyEditorHtml();
    }

    webviewView.webview.onDidReceiveMessage(async (message) => {
      if (message?.type === 'revealField') {
        const fieldId = typeof message.fieldId === 'string' ? message.fieldId : '';
        if (fieldId.length > 0) {
          void this.revealActiveEntityField(fieldId);
        }
        return;
      }

      if (message?.type === 'updateField') {
        const fieldId = typeof message.fieldId === 'string' ? message.fieldId : '';
        const value = message.value;
        if (fieldId.length === 0) {
          return;
        }

        if (!this.activeEntityKind) {
          return;
        }

        if (this.activeEntityIsDraft && this.activeEntityDraft) {
          this.updateDraftEntityFieldValue(this.activeEntityKind, this.activeEntityDraft, fieldId, value);
          return;
        }

        if (!this.activeEntityMap) {
          return;
        }

        if (this.updateEntityFieldValue(this.activeEntityKind, this.activeEntityMap, fieldId, value)) {
          this.persistYaml();
        }
        return;
      }

      if (message?.type !== 'save') {
        return;
      }

      if (!this.activeEntityKind) {
        return;
      }

      const payload = message.payload as EntityEditorPayload;
      const parseResult = this.validatePayloadForKind(this.activeEntityKind, payload);
      if (!parseResult.ok) {
        vscode.window.showErrorMessage(parseResult.error);
        return;
      }

      if (this.activeEntityIsDraft) {
        const insertedEntity = this.insertEntityFromPayload(this.activeEntityKind, payload);
        this.activeEntityMap = insertedEntity;
        this.activeEntityDraft = undefined;
        this.activeEntityIsDraft = false;
        this.activeEntityLabel = this.getEntityLabel(this.activeEntityKind, insertedEntity, 'New Item');
        this.refresh();

        const insertedSelection = this.getNodeSelectionOffsets(insertedEntity);
        if (insertedSelection) {
          await this.revealOffsets(insertedSelection[0], insertedSelection[1], true);
        }

        vscode.window.showInformationMessage('Item inserted.');
        return;
      }

      if (!this.activeEntityMap) {
        return;
      }

      this.updateEntityNode(this.activeEntityKind, this.activeEntityMap, payload);
      this.persistYaml();
      this.refresh();
      vscode.window.showInformationMessage('Item updated.');
    });
  }


  setTreeView(treeView: vscode.TreeView<YAMLTreeItem>): void {
    this.treeView = treeView;
  }

  getTreeItem(element: YAMLTreeItem): vscode.TreeItem {
    return element;
  }

  getParent(element: YAMLTreeItem): YAMLTreeItem | undefined {
    if (!element.modelFilePath) {
      return undefined;
    }
    // Walk the cached tree items to find the parent model node
    for (const [, treeItem] of this.treeItemsByModelPath) {
      if (treeItem.children.includes(element)) {
        return treeItem;
      }
    }
    return undefined;
  }

  getChildren(element?: YAMLTreeItem): Thenable<YAMLTreeItem[]> {
    console.log('getChildren');
    if (!this.workspaceRoot) {
      vscode.window.showInformationMessage('No dependency in empty workspace');
      return Promise.resolve([]);
    }

    if (element) {
      if (element.children.length > 0) {
        return Promise.resolve(element.children);
      }

      if (element.nodeKind === 'model' && element.modelNode) {
        const children = this.getModelNodeChildren(element.modelNode);
        element.children = children;
        return Promise.resolve(children);
      }

      if (element.item && YAML.isMap(element.item)) {
        const children = this.getMapChildren(element.item);
        children.forEach(c => { c.modelFilePath = element.modelFilePath; });
        element.children = children;
        return Promise.resolve(children);
      }

      if (element.item && YAML.isSeq(element.item)) {
        const children = this.getSeqChildren(element.item);
        children.forEach(c => { c.modelFilePath = element.modelFilePath; });
        element.children = children;
        return Promise.resolve(children);
      }

      return Promise.resolve([]);
    }
    else {
      // we're at the root, so return the top-level hierarchy
      return Promise.resolve(this.getRootItems());
    }
  }

  private getRootItems(): YAMLTreeItem[] {
    if (this.modelTree) {
      return [this.createModelTreeItem(this.modelTree, true)];
    }
    return [];
  }

  private getNamedSequenceItems(seq: YAML.YAMLSeq | null, fallbackLabel: string, category: TreeItemCategory): YAMLTreeItem[] {
    if (!seq || !Array.isArray(seq.items)) {
      return [];
    }

    return seq.items.map((node, index) => {
      const label = this.getNodeLabel(node, `${fallbackLabel} ${index + 1}`);
      return new YAMLTreeItem(label, vscode.TreeItemCollapsibleState.None, [], node, YAML.isMap(node) ? 'map' : 'scalar', category);
    });
  }

  // -- Model hierarchy methods --

  private buildFullHierarchy(startFilePath: string): void {
    const rootPath = this.discoverRoot(startFilePath);
    this.modelTree = this.loadModelNode(rootPath);
    this.modelsByPath.clear();
    this.treeItemsByModelPath.clear();
    this.indexModelNodes(this.modelTree);
  }

  private discoverRoot(filePath: string): string {
    const visited = new Set<string>();
    let currentPath = filePath;

    while (true) {
      if (visited.has(currentPath)) {
        break;
      }
      visited.add(currentPath);

      let parsed: ParsedModel;
      try {
        parsed = this.parseAndCache(currentPath);
      } catch {
        break;
      }

      const parentId = parsed.doc.get('parent');

      if (typeof parentId !== 'string' || parentId.length === 0) {
        return currentPath;
      }

      const parentPath = this.resolveParentPath(currentPath, parentId);
      if (!fs.existsSync(parentPath)) {
        return currentPath;
      }

      currentPath = parentPath;
    }

    return currentPath;
  }

  private loadModelNode(filePath: string): ModelNode {
    const parsed = this.parseAndCache(filePath);
    const { source, doc, model } = parsed;

    const children: ModelNode[] = [];
    const childrenList = doc.get('children');
    if (childrenList && YAML.isSeq(childrenList)) {
      for (const childItem of childrenList.items) {
        if (YAML.isMap(childItem)) {
          const refId = childItem.get('REFID');
          if (typeof refId === 'string') {
            const childPath = this.resolveChildPath(filePath, refId);
            if (fs.existsSync(childPath)) {
              children.push(this.loadModelNode(childPath));
            }
          }
        }
      }
    }

    return {
      id: model.ID || path.basename(filePath, '.yaml'),
      title: model.title || model.ID || path.basename(filePath, '.yaml'),
      filePath,
      doc,
      model,
      source,
      children
    };
  }

  private resolveParentPath(childFilePath: string, parentId: string): string {
    const childDir = path.dirname(childFilePath);
    const parentDir = path.dirname(childDir);
    return path.join(parentDir, parentId + '.yaml');
  }

  private resolveChildPath(parentFilePath: string, childRefId: string): string {
    const parentDir = path.dirname(parentFilePath);
    return path.join(parentDir, childRefId, childRefId + '.yaml');
  }

  private indexModelNodes(node: ModelNode): void {
    this.modelsByPath.set(node.filePath, node);
    for (const child of node.children) {
      this.indexModelNodes(child);
    }
  }

  private switchToModel(modelNode: ModelNode): void {
    this.workspaceRoot = modelNode.filePath;
    this.tmDoc = modelNode.doc;
    this.model = modelNode.model;
    this.tmSource = modelNode.source;
  }

  private syncModelNode(): void {
    const modelNode = this.modelsByPath.get(this.workspaceRoot);
    if (modelNode && this.tmDoc && this.model && this.tmSource !== undefined) {
      modelNode.doc = this.tmDoc;
      modelNode.model = this.model;
      modelNode.source = this.tmSource;
    }
  }

  private createModelTreeItem(modelNode: ModelNode, expanded = false): YAMLTreeItem {
    const item = new YAMLTreeItem(
      modelNode.title,
      expanded ? vscode.TreeItemCollapsibleState.Expanded : vscode.TreeItemCollapsibleState.Collapsed,
      [],
      undefined,
      'model',
      'model'
    );
    item.modelNode = modelNode;
    item.modelFilePath = modelNode.filePath;
    item.id = 'model:' + modelNode.filePath;
    // Use the model's directory as resourceUri so VS Code resolves the
    // file icon theme folder icon (including open/closed states)
    item.resourceUri = vscode.Uri.file(path.dirname(modelNode.filePath));
    item.iconPath = undefined;
    this.treeItemsByModelPath.set(modelNode.filePath, item);
    if (this.pendingRevealPath === modelNode.filePath) {
      this.pendingRevealPath = undefined;
      if (this.treeView) {
        this.suppressRevealSelection = true;
        void this.treeView.reveal(item, { select: true, focus: false, expand: true }).then(
          () => { this.suppressRevealSelection = false; },
          () => { this.suppressRevealSelection = false; }
        );
      }
    }
    return item;
  }

  private getModelNodeChildren(modelNode: ModelNode): YAMLTreeItem[] {
    const doc = modelNode.doc;
    const result: YAMLTreeItem[] = [];

    const threats = doc.get('threats') as YAML.YAMLSeq | null;
    const scope = doc.get('scope') as YAML.YAMLMap | null;
    const secObjectives = scope ? scope.get('securityObjectives') as YAML.YAMLSeq | null : null;
    const attackers = scope ? scope.get('attackers') as YAML.YAMLSeq | null : null;

    const addSection = (label: string, seq: YAML.YAMLSeq | null, fallback: string, category: TreeItemCategory) => {
      const items = this.getNamedSequenceItems(seq, fallback, category);
      items.forEach(i => { i.modelFilePath = modelNode.filePath; });
      const root = new YAMLTreeItem(label, vscode.TreeItemCollapsibleState.Collapsed, items, seq ?? undefined, 'root', category);
      root.modelFilePath = modelNode.filePath;
      root.id = modelNode.filePath + ':' + category;
      result.push(root);
    };

    addSection('Threats', threats, 'Threat', 'threat');
    addSection('Security Objectives', secObjectives, 'Security Objective', 'securityObjective');
    addSection('Attackers', attackers, 'Attacker', 'attacker');

    for (const child of modelNode.children) {
      result.push(this.createModelTreeItem(child));
    }

    return result;
  }

  private getNodeLabel(node: unknown, fallback: string): string {
    if (!node || !YAML.isMap(node)) {
      return fallback;
    }

    const candidateKeys = ['title', 'name', 'ID', 'id', 'REFID', 'refid'];
    for (const key of candidateKeys) {
      const value = node.get(key);
      if (typeof value === 'string' && value.trim().length > 0) {
        return value;
      }
    }

    return fallback;
  }

  setActiveFile(filePath: string): void {
    if (filePath === this.workspaceRoot) {
      return;
    }

    // Check if this file is already in the current hierarchy
    const existingNode = this.modelsByPath.get(filePath);
    if (existingNode) {
      this.switchToModel(existingNode);
      this.openModelEditorForFile(filePath);
      this.revealModelTreeItem(filePath);
      return;
    }

    // Not in current hierarchy — check if it's a threat model YAML
    if (!this.isThreatModelFile(filePath)) {
      this.clear();
      return;
    }

    // Different hierarchy — rebuild
    this.clearActiveEntity();
    this.treeItemsByModelPath.clear();
    this.buildFullHierarchy(filePath);

    const activeNode = this.modelsByPath.get(filePath);
    if (activeNode) {
      this.switchToModel(activeNode);
    } else {
      this.workspaceRoot = filePath;
      this.reloadFromSource();
    }

    this._onDidChangeTreeData.fire();
    this.openModelEditorForFile(filePath);
    this.revealModelTreeItem(filePath);
  }

  private openModelEditorForFile(filePath: string): void {
    const modelNode = this.modelsByPath.get(filePath);
    if (!modelNode) {
      return;
    }
    const contents = modelNode.doc.contents;
    if (!contents || !YAML.isMap(contents)) {
      return;
    }
    this.activeEntityKind = 'model';
    this.activeEntityMap = contents;
    this.activeEntityDraft = undefined;
    this.activeEntityIsDraft = false;
    this.activeEntityIndex = undefined;
    this.activeEntityId = undefined;
    this.activeEntityLabel = modelNode.title || 'Model';
    this.renderActiveEditor();
  }

  private isThreatModelFile(filePath: string): boolean {
    try {
      const parsed = this.parseAndCache(filePath);
      const contents = parsed.doc.contents;
      if (!contents || !YAML.isMap(contents)) {
        return false;
      }
      // A threat model file has an ID field at the root
      const id = (contents as YAML.YAMLMap).get('ID');
      return typeof id === 'string' && id.length > 0;
    } catch {
      return false;
    }
  }

  clear(): void {
    this.modelTree = null;
    this.modelsByPath.clear();
    this.treeItemsByModelPath.clear();
    this.clearActiveEntity();
    this._onDidChangeTreeData.fire();
    this.renderActiveEditor();
  }

  private clearActiveEntity(): void {
    this.activeEntityKind = undefined;
    this.activeEntityMap = undefined;
    this.activeEntityDraft = undefined;
    this.activeEntityIsDraft = false;
    this.activeEntityIndex = undefined;
    this.activeEntityId = undefined;
  }

  isEditorSwitchSuppressed(): boolean {
    return this.editorSwitchSuppressCount > 0;
  }

  isRevealSelection(): boolean {
    return this.suppressRevealSelection;
  }

  private revealModelTreeItem(filePath: string): void {
    if (!this.treeView) {
      return;
    }
    const treeItem = this.treeItemsByModelPath.get(filePath);
    if (treeItem) {
      this.suppressRevealSelection = true;
      void this.treeView.reveal(treeItem, { select: true, focus: false, expand: true }).then(
        () => { this.suppressRevealSelection = false; },
        () => { this.suppressRevealSelection = false; }
      );
    } else {
      // Tree items haven't been created yet; defer the reveal
      this.pendingRevealPath = filePath;
    }
  }

  getActiveFile(): string {
    return this.workspaceRoot;
  }

  refresh(): void {
    console.log('refresh');
    this.parsedCache.clear();
    const activePath = this.workspaceRoot;
    this.buildFullHierarchy(activePath);

    const activeNode = this.modelsByPath.get(activePath);
    if (activeNode) {
      this.switchToModel(activeNode);
    } else {
      this.reloadFromSource();
    }

    if (!this.activeEntityIsDraft) {
      this.syncActiveEntityFromSource();
    }

    this._onDidChangeTreeData.fire();
    this.renderActiveEditor();
  }

  handleYamlSourceChanged(): void {
    if (this.suppressNextExternalChange) {
      this.suppressNextExternalChange = false;
      return;
    }

    this.reloadFromSource();
    this.syncModelNode();
    this._onDidChangeTreeData.fire();
    this.renderActiveEditor();
  }

  async revealTreeItem(item: YAMLTreeItem): Promise<void> {
    // For model nodes, open the model's YAML file
    if (item.nodeKind === 'model' && item.modelNode) {
      const uri = vscode.Uri.file(item.modelNode.filePath);
      this.editorSwitchSuppressCount++;
      try {
        await vscode.window.showTextDocument(uri, { preview: false });
      } finally {
        this.editorSwitchSuppressCount--;
      }
      return;
    }

    // Switch to the correct model for source text / file context
    if (item.modelFilePath && item.modelFilePath !== this.workspaceRoot) {
      const modelNode = this.modelsByPath.get(item.modelFilePath);
      if (modelNode) {
        this.switchToModel(modelNode);
      }
    }

    const startOffset = this.getNodeStartOffset(item.item);
    if (startOffset === undefined) {
      vscode.window.showInformationMessage('No source location for this tree item.');
      return;
    }

    await this.revealOffset(startOffset);
  }

  private async revealActiveEntityField(fieldId: string): Promise<void> {
    if (!this.activeEntityMap || !this.activeEntityKind) {
      return;
    }

    const fieldNode = this.getEntityFieldNode(this.activeEntityKind, this.activeEntityMap, fieldId);
    const selectionOffsets = this.getNodeSelectionOffsets(fieldNode);
    if (!selectionOffsets) {
      return;
    }

    await this.revealOffsets(selectionOffsets[0], selectionOffsets[1], true);
  }

  private async revealOffset(startOffset: number): Promise<void> {
    await this.revealOffsets(startOffset, startOffset, false);
  }

  private async revealOffsets(startOffset: number, endOffset: number, preserveFocus: boolean): Promise<void> {
    const position = this.offsetToPosition(startOffset);
    const endPosition = this.offsetToPosition(endOffset);
    const selectionRange = new vscode.Range(position, endPosition);
    const uri = vscode.Uri.file(this.workspaceRoot);
    this.editorSwitchSuppressCount++;
    try {
      const editor = await vscode.window.showTextDocument(uri, {
        preview: false,
        preserveFocus,
        selection: selectionRange
      });
      editor.revealRange(selectionRange, vscode.TextEditorRevealType.InCenter);
    } finally {
      this.editorSwitchSuppressCount--;
    }
  }

  private getEntityFieldNode(kind: EditorEntityKind, entityMap: YAML.YAMLMap, fieldId: string): unknown {
    if (kind === 'threat' && fieldId === 'cvssVector') {
      const cvssNode = this.getMapValueNode(entityMap, 'CVSS');
      if (cvssNode && YAML.isMap(cvssNode)) {
        return this.getMapValueNode(cvssNode, 'vector') ?? cvssNode;
      }
      return cvssNode;
    }

    const keyMapByKind: Record<EditorEntityKind, Record<string, string>> = {
      threat: {
        ID: 'ID',
        title: 'title',
        attack: 'attack',
        impactDesc: 'impactDesc',
        threatType: 'threatType',
        fullyMitigated: 'fullyMitigated',
        public: 'public',
      },
      securityObjective: {
        ID: 'ID',
        title: 'title',
        description: 'description'
      },
      attacker: {
        ID: 'ID',
        name: 'name',
        title: 'title',
        description: 'description'
      },
      model: {
        ID: 'ID',
        title: 'title',
        version: 'version',
        analysis: 'analysis'
      }
    };

    const keyMap = keyMapByKind[kind];
    const yamlKey = keyMap[fieldId];
    if (!yamlKey) {
      return undefined;
    }

    return this.getMapValueNode(entityMap, yamlKey);
  }

  private getMapValueNode(map: YAML.YAMLMap, key: string): unknown {
    for (const pair of map.items) {
      if (String(pair.key) === key) {
        return pair.value;
      }
    }
    return undefined;
  }

  isEditableItem(item: YAMLTreeItem): boolean {
    if (item.category === 'model' && item.modelNode != null) {
      return true;
    }
    return (item.category === 'threat' || item.category === 'securityObjective' || item.category === 'attacker') && this.getEntityMap(item.item) !== null;
  }

  openEntityEditor(item: YAMLTreeItem): void {
    if (!this.isEditableItem(item)) {
      return;
    }

    // Model node: the entity map is the document root contents
    if (item.category === 'model' && item.modelNode) {
      this.switchToModel(item.modelNode);
      const contents = item.modelNode.doc.contents;
      if (!contents || !YAML.isMap(contents)) {
        return;
      }
      this.activeEntityKind = 'model';
      this.activeEntityMap = contents;
      this.activeEntityDraft = undefined;
      this.activeEntityIsDraft = false;
      this.activeEntityIndex = undefined;
      this.activeEntityId = undefined;
      this.activeEntityLabel = item.modelNode.title || 'Model';
      if (this.editorView) {
        this.editorView.webview.html = this.getEntityEditorHtml('model', contents, this.activeEntityLabel, this.pendingFocusField);
        this.pendingFocusField = undefined;
      }
      this.editorSwitchSuppressCount++;
      void vscode.commands.executeCommand('tmexpThreatEditorView.focus').then(
        () => { this.editorSwitchSuppressCount--; },
        () => { this.editorSwitchSuppressCount--; }
      );
      return;
    }

    // Switch to the correct model if needed
    if (item.modelFilePath && item.modelFilePath !== this.workspaceRoot) {
      const modelNode = this.modelsByPath.get(item.modelFilePath);
      if (modelNode) {
        this.switchToModel(modelNode);
      }
    }

    const entityKind = item.category as EditorEntityKind;
    const entityMap = this.getEntityMap(item.item);
    if (!entityMap) {
      return;
    }

    this.activeEntityKind = entityKind;
    this.activeEntityMap = entityMap;
    this.activeEntityDraft = undefined;
    this.activeEntityIsDraft = false;
    this.activeEntityIndex = this.findEntityIndexByMap(entityKind, entityMap);
    this.activeEntityId = this.getEntityId(entityKind, entityMap);
    this.activeEntityLabel = item.label;
    if (this.editorView) {
      this.editorView.webview.html = this.getEntityEditorHtml(entityKind, entityMap, item.label, this.pendingFocusField);
      this.pendingFocusField = undefined;
    }

    this.editorSwitchSuppressCount++;
    void vscode.commands.executeCommand('tmexpThreatEditorView.focus').then(
      () => { this.editorSwitchSuppressCount--; },
      () => { this.editorSwitchSuppressCount--; }
    );
  }

  createNewThreat(): void {
    this.createNewEntity('threat');
  }

  createNewEntity(kind: EditorEntityKind, modelFilePath?: string): void {
    if (modelFilePath) {
      const modelNode = this.modelsByPath.get(modelFilePath);
      if (modelNode) {
        this.switchToModel(modelNode);
      }
    }

    const labels: Record<EditorEntityKind, string> = {
      threat: 'New Threat',
      securityObjective: 'New Security Objective',
      attacker: 'New Attacker',
      model: 'New Model'
    };

    this.activeEntityKind = kind;
    this.activeEntityMap = undefined;
    this.activeEntityDraft = this.getDefaultPayloadForKind(kind);
    this.activeEntityIsDraft = true;
    this.activeEntityIndex = undefined;
    this.activeEntityId = undefined;
    this.activeEntityLabel = labels[kind];
    this.pendingFocusField = 'ID';

    if (this.editorView && this.activeEntityKind && this.activeEntityDraft) {
      this.editorView.webview.html = this.getEntityEditorHtmlFromPayload(this.activeEntityKind, this.activeEntityDraft, this.activeEntityLabel, this.pendingFocusField);
      this.pendingFocusField = undefined;
    }

    void vscode.commands.executeCommand('tmexpThreatEditorView.focus');
  }

  private getMapChildren(map: YAML.YAMLMap): YAMLTreeItem[] {
    return map.items.map((pair) => {
      const key = String(pair.key);
      return this.createTreeItemForValue(key, pair.value);
    });
  }

  private getSeqChildren(seq: YAML.YAMLSeq): YAMLTreeItem[] {
    return seq.items.map((item, index) => {
      return this.createTreeItemForValue(`[${index}]`, item);
    });
  }

  private createTreeItemForValue(label: string, value: unknown): YAMLTreeItem {
    if (value && YAML.isMap(value)) {
      return new YAMLTreeItem(label, vscode.TreeItemCollapsibleState.Collapsed, [], value, 'map', 'other');
    }

    if (value && YAML.isSeq(value)) {
      return new YAMLTreeItem(label, vscode.TreeItemCollapsibleState.Collapsed, [], value, 'seq', 'other');
    }

    if (value && YAML.isScalar(value)) {
      return new YAMLTreeItem(
        `${label}: ${this.formatValue(value.value)}`,
        vscode.TreeItemCollapsibleState.None,
        [],
        value,
        'scalar',
        'other'
      );
    }

    return new YAMLTreeItem(
      `${label}: ${this.formatValue(value)}`,
      vscode.TreeItemCollapsibleState.None,
      [],
      value,
      'scalar',
      'other'
    );
  }

  private getEntityEditorHtml(kind: EditorEntityKind, entityMap: YAML.YAMLMap, label: string, focusField?: string): string {
    return this.getEntityEditorHtmlFromPayload(kind, this.extractEntityEditorData(kind, entityMap), label, focusField);
  }

  private getEntityEditorHtmlFromPayload(kind: EditorEntityKind, payload: EntityEditorPayload, label: string, focusField?: string): string {
    return renderEntityEditorHtml({
      kind,
      label,
      payload,
      focusField
    });
  }

  private extractEntityEditorData(kind: EditorEntityKind, entityMap: YAML.YAMLMap): EntityEditorPayload {
    if (kind === 'threat') {
      return this.extractThreatEditorData(entityMap);
    }

    if (kind === 'model') {
      const payload: ThreatModelEditorPayload = {
        ID: this.getMapString(entityMap, 'ID'),
        title: this.getMapString(entityMap, 'title'),
        version: this.getMapString(entityMap, 'version'),
        analysis: this.getMapString(entityMap, 'analysis')
      };
      return payload;
    }

    if (kind === 'securityObjective') {
      const payload: SecurityObjectiveEditorPayload = {
        ID: this.getMapString(entityMap, 'ID'),
        title: this.getMapString(entityMap, 'title'),
        description: this.getMapString(entityMap, 'description')
      };
      return payload;
    }

    const attackerPayload: AttackerEditorPayload = {
      ID: this.getMapString(entityMap, 'ID'),
      name: this.getMapString(entityMap, 'name'),
      title: this.getMapString(entityMap, 'title'),
      description: this.getMapString(entityMap, 'description')
    };
    return attackerPayload;
  }

  private extractThreatEditorData(threatMap: YAML.YAMLMap): ThreatEditorPayload {
    const cvss = threatMap.get('CVSS');
    const cvssVector = (cvss && YAML.isMap(cvss) && typeof cvss.get('vector') === 'string') ? cvss.get('vector') as string : '';

    return {
      ID: this.getMapString(threatMap, 'ID'),
      title: this.getMapString(threatMap, 'title'),
      attack: this.getMapString(threatMap, 'attack'),
      impactDesc: this.getMapString(threatMap, 'impactDesc'),
      fullyMitigated: this.getMapBoolean(threatMap, 'fullyMitigated'),
      cvssVector,
      threatType: this.getMapString(threatMap, 'threatType'),
      public: this.getMapBoolean(threatMap, 'public'),
    };
  }

  private getMapString(map: YAML.YAMLMap, key: string): string {
    const value = map.get(key);
    return typeof value === 'string' ? value : '';
  }

  private getMapBoolean(map: YAML.YAMLMap, key: string): boolean {
    const value = map.get(key);
    return typeof value === 'boolean' ? value : false;
  }

  private getNodeAsJs(node: unknown): unknown {
    if (!node) {
      return undefined;
    }

    if (YAML.isMap(node) || YAML.isSeq(node) || YAML.isScalar(node)) {
      return node.toJSON();
    }

    return node;
  }

  private getEntityMap(node: unknown): YAML.YAMLMap | null {
    if (node && YAML.isMap(node)) {
      return node;
    }

    return null;
  }

  private updateEntityNode(kind: EditorEntityKind, entityMap: YAML.YAMLMap, payload: EntityEditorPayload): void {
    if (kind === 'model') {
      const p = payload as ThreatModelEditorPayload;
      entityMap.set('ID', p.ID);
      entityMap.set('title', p.title);
      entityMap.set('version', p.version);
      entityMap.set('analysis', p.analysis);
      return;
    }

    if (kind === 'threat') {
      const p = payload as ThreatEditorPayload;
      entityMap.set('ID', p.ID);
      entityMap.set('title', p.title);
      entityMap.set('attack', p.attack);
      entityMap.set('impactDesc', p.impactDesc);
      entityMap.set('fullyMitigated', p.fullyMitigated);
      entityMap.set('threatType', p.threatType);
      entityMap.set('public', p.public);

      const existingCvss = entityMap.get('CVSS');
      if (existingCvss && YAML.isMap(existingCvss)) {
        existingCvss.set('vector', p.cvssVector);
      } else {
        entityMap.set('CVSS', { vector: p.cvssVector });
      }
      return;
    }

    if (kind === 'securityObjective') {
      const p = payload as SecurityObjectiveEditorPayload;
      entityMap.set('ID', p.ID);
      entityMap.set('title', p.title);
      entityMap.set('description', p.description);
      return;
    }

    const p = payload as AttackerEditorPayload;
    entityMap.set('ID', p.ID);
    entityMap.set('name', p.name);
    entityMap.set('title', p.title);
    entityMap.set('description', p.description);
  }

  private getDefaultPayloadForKind(kind: EditorEntityKind): EntityEditorPayload {
    if (kind === 'threat') {
      return this.getDefaultThreatPayload();
    }
    if (kind === 'securityObjective') {
      return { ID: '', title: 'New Security Objective', description: '' } as SecurityObjectiveEditorPayload;
    }
    if (kind === 'model') {
      return { ID: '', title: '', version: '', analysis: '' } as ThreatModelEditorPayload;
    }
    return { ID: '', name: '', title: 'New Attacker', description: '' } as AttackerEditorPayload;
  }

  private getDefaultThreatPayload(): ThreatEditorPayload {
    return {
      ID: '',
      title: 'New Threat',
      attack: '',
      impactDesc: '',
      fullyMitigated: false,
      cvssVector: '',
      threatType: '',
      public: false
    };
  }

  private updateDraftEntityFieldValue(kind: EditorEntityKind, payload: EntityEditorPayload, fieldId: string, rawValue: unknown): void {
    const value = rawValue;
    if (kind === 'model') {
      const p = payload as ThreatModelEditorPayload;
      if (fieldId === 'ID' || fieldId === 'title' || fieldId === 'version' || fieldId === 'analysis') {
        p[fieldId] = typeof value === 'string' ? value : String(value ?? '');
      }
      return;
    }

    if (kind === 'threat') {
      const p = payload as ThreatEditorPayload;
      if (fieldId === 'ID' || fieldId === 'title' || fieldId === 'attack' || fieldId === 'impactDesc' || fieldId === 'threatType' || fieldId === 'cvssVector') {
        p[fieldId] = typeof value === 'string' ? value : String(value ?? '');
        return;
      }
      if (fieldId === 'fullyMitigated' || fieldId === 'public') {
        p[fieldId] = Boolean(value);
      }
      return;
    }

    if (kind === 'securityObjective') {
      const p = payload as SecurityObjectiveEditorPayload;
      if (fieldId === 'ID' || fieldId === 'title' || fieldId === 'description') {
        p[fieldId] = typeof value === 'string' ? value : String(value ?? '');
      }
      return;
    }

    const p = payload as AttackerEditorPayload;
    if (fieldId === 'ID' || fieldId === 'name' || fieldId === 'title' || fieldId === 'description') {
      p[fieldId] = typeof value === 'string' ? value : String(value ?? '');
    }
  }

  private insertEntityFromPayload(kind: EditorEntityKind, payload: EntityEditorPayload): YAML.YAMLMap {
    let collection = this.getCollectionByKind(kind, true);
    if (!collection) {
      collection = new YAML.YAMLSeq();
      this.setCollectionByKind(kind, collection);
    }

    const insertionIndex = collection.items.length;
    const newEntity = new YAML.YAMLMap();
    this.updateEntityNode(kind, newEntity, payload);
    collection.add(newEntity);
    this.persistYaml();

    this.invalidateCache(this.workspaceRoot);
    const reloaded = this.parseAndCache(this.workspaceRoot);
    const reloadedCollection = this.getCollectionByKindFromDoc(kind, reloaded.doc);
    if (reloadedCollection && reloadedCollection.items[insertionIndex] && YAML.isMap(reloadedCollection.items[insertionIndex])) {
      this.tmDoc = reloaded.doc;
      this.model = reloaded.model;
      this.tmSource = reloaded.source;
      this.syncModelNode();
      return reloadedCollection.items[insertionIndex] as YAML.YAMLMap;
    }

    return newEntity;
  }

  private updateEntityFieldValue(kind: EditorEntityKind, entityMap: YAML.YAMLMap, fieldId: string, rawValue: unknown): boolean {
    if (kind === 'model') {
      if (fieldId === 'ID' || fieldId === 'title' || fieldId === 'version' || fieldId === 'analysis') {
        const value = typeof rawValue === 'string' ? rawValue : String(rawValue ?? '');
        entityMap.set(fieldId, value);
        return true;
      }
      return false;
    }

    if (kind === 'securityObjective') {
      if (fieldId === 'ID' || fieldId === 'title' || fieldId === 'description') {
        const value = typeof rawValue === 'string' ? rawValue : String(rawValue ?? '');
        entityMap.set(fieldId, value);
        return true;
      }
      return false;
    }

    if (kind === 'attacker') {
      if (fieldId === 'ID' || fieldId === 'name' || fieldId === 'title' || fieldId === 'description') {
        const value = typeof rawValue === 'string' ? rawValue : String(rawValue ?? '');
        entityMap.set(fieldId, value);
        return true;
      }
      return false;
    }

    const stringFields: Record<string, string> = {
      ID: 'ID',
      title: 'title',
      attack: 'attack',
      impactDesc: 'impactDesc',
      threatType: 'threatType'
    };

    const booleanFields: Record<string, string> = {
      fullyMitigated: 'fullyMitigated',
      public: 'public'
    };

    const stringKey = stringFields[fieldId];
    if (stringKey) {
      const value = typeof rawValue === 'string' ? rawValue : String(rawValue ?? '');
      entityMap.set(stringKey, value);
      return true;
    }

    const booleanKey = booleanFields[fieldId];
    if (booleanKey) {
      entityMap.set(booleanKey, Boolean(rawValue));
      return true;
    }

    if (fieldId === 'cvssVector') {
      const vectorValue = typeof rawValue === 'string' ? rawValue : String(rawValue ?? '');
      const existingCvss = entityMap.get('CVSS');
      if (existingCvss && YAML.isMap(existingCvss)) {
        existingCvss.set('vector', vectorValue);
      } else {
        entityMap.set('CVSS', { vector: vectorValue });
      }
      return true;
    }

    return false;
  }

  private validatePayloadForKind(kind: EditorEntityKind, payload: EntityEditorPayload): { ok: true } | { ok: false; error: string } {
    if (kind !== 'threat') {
      return { ok: true };
    }

    return { ok: true };
  }

  private getEntityLabel(kind: EditorEntityKind, entityMap: YAML.YAMLMap, fallback: string): string {
    if (kind === 'attacker') {
      const name = entityMap.get('name');
      if (typeof name === 'string' && name.trim().length > 0) {
        return name;
      }
    }

    return this.getNodeLabel(entityMap, fallback);
  }

  private getCollectionByKind(kind: EditorEntityKind, createIfMissing: boolean): YAML.YAMLSeq | null {
    if (!this.tmDoc) {
      return null;
    }
    return this.getCollectionByKindFromDoc(kind, this.tmDoc, createIfMissing);
  }

  private getCollectionByKindFromDoc(kind: EditorEntityKind, doc: YAML.Document, createIfMissing = false): YAML.YAMLSeq | null {
    if (kind === 'threat') {
      const threats = doc.get('threats') as YAML.YAMLSeq | null;
      if (threats && YAML.isSeq(threats)) {
        return threats;
      }
      if (createIfMissing) {
        const seq = new YAML.YAMLSeq();
        doc.set('threats', seq);
        return seq;
      }
      return null;
    }

    let scope = doc.get('scope') as YAML.YAMLMap | null;
    if (!scope || !YAML.isMap(scope)) {
      if (!createIfMissing) {
        return null;
      }
      scope = new YAML.YAMLMap();
      doc.set('scope', scope);
    }

    const key = kind === 'securityObjective' ? 'securityObjectives' : 'attackers';
    const existing = scope.get(key) as YAML.YAMLSeq | null;
    if (existing && YAML.isSeq(existing)) {
      return existing;
    }

    if (!createIfMissing) {
      return null;
    }

    const seq = new YAML.YAMLSeq();
    scope.set(key, seq);
    return seq;
  }

  private setCollectionByKind(kind: EditorEntityKind, seq: YAML.YAMLSeq): void {
    if (!this.tmDoc) {
      return;
    }
    if (kind === 'threat') {
      this.tmDoc.set('threats', seq);
      return;
    }

    let scope = this.tmDoc.get('scope') as YAML.YAMLMap | null;
    if (!scope || !YAML.isMap(scope)) {
      scope = new YAML.YAMLMap();
      this.tmDoc.set('scope', scope);
    }

    scope.set(kind === 'securityObjective' ? 'securityObjectives' : 'attackers', seq);
  }

  private persistYaml(): void {
    if (!this.tmDoc) {
      return;
    }
    this.suppressNextExternalChange = true;
    const serialized = String(this.tmDoc);
    fs.writeFileSync(this.workspaceRoot, serialized, 'utf8');
    // Re-parse from the serialized string to get fresh source tokens/ranges
    const doc = tmlib.parseYAMLFromString(serialized);
    const model = tmlib.parseThreatModelFromString(serialized);
    this.tmSource = serialized;
    this.tmDoc = doc;
    this.model = model;
    this.parsedCache.set(this.workspaceRoot, { source: serialized, doc, model });
    this.syncModelNode();
  }

  private reloadFromSource(): void {
    this.invalidateCache(this.workspaceRoot);
    const parsed = this.parseAndCache(this.workspaceRoot);
    this.tmSource = parsed.source;
    this.model = parsed.model;
    this.tmDoc = parsed.doc;

    if (!this.activeEntityIsDraft) {
      this.syncActiveEntityFromSource();
    }
  }

  private parseAndCache(filePath: string): ParsedModel {
    const cached = this.parsedCache.get(filePath);
    if (cached) {
      return cached;
    }
    const source = fs.readFileSync(filePath, 'utf8');
    const doc = tmlib.parseYAMLFromString(source);
    const model = tmlib.parseThreatModelFromString(source);
    const entry: ParsedModel = { source, doc, model };
    this.parsedCache.set(filePath, entry);
    return entry;
  }

  private invalidateCache(filePath: string): void {
    this.parsedCache.delete(filePath);
  }

  private syncActiveEntityFromSource(): void {
    if (!this.activeEntityKind || !this.tmDoc) {
      return;
    }

    if (this.activeEntityKind === 'model') {
      const contents = this.tmDoc.contents;
      if (contents && YAML.isMap(contents)) {
        this.activeEntityMap = contents;
      }
      return;
    }

    const collection = this.getCollectionByKind(this.activeEntityKind, false);
    if (!collection || !Array.isArray(collection.items)) {
      this.activeEntityMap = undefined;
      this.activeEntityIndex = undefined;
      this.activeEntityId = undefined;
      return;
    }

    let matchedIndex: number | undefined;
    let matchedMap: YAML.YAMLMap | undefined;

    if (this.activeEntityId && this.activeEntityId.length > 0) {
      collection.items.forEach((node, index) => {
        if (!matchedMap && YAML.isMap(node) && this.getEntityId(this.activeEntityKind as EditorEntityKind, node) === this.activeEntityId) {
          matchedMap = node;
          matchedIndex = index;
        }
      });
    }

    if (!matchedMap && typeof this.activeEntityIndex === 'number') {
      const byIndex = collection.items[this.activeEntityIndex];
      if (byIndex && YAML.isMap(byIndex)) {
        matchedMap = byIndex;
        matchedIndex = this.activeEntityIndex;
      }
    }

    this.activeEntityMap = matchedMap;
    this.activeEntityIndex = matchedIndex;
    this.activeEntityId = matchedMap ? this.getEntityId(this.activeEntityKind, matchedMap) : undefined;
    if (matchedMap) {
      this.activeEntityLabel = this.getEntityLabel(this.activeEntityKind, matchedMap, 'Item');
    }
  }

  private renderActiveEditor(): void {
    if (!this.editorView) {
      return;
    }

    if (this.activeEntityKind && this.activeEntityMap) {
      this.editorView.webview.html = this.getEntityEditorHtml(this.activeEntityKind, this.activeEntityMap, this.activeEntityLabel, this.pendingFocusField);
      this.pendingFocusField = undefined;
      return;
    }

    if (this.activeEntityKind && this.activeEntityDraft) {
      this.editorView.webview.html = this.getEntityEditorHtmlFromPayload(this.activeEntityKind, this.activeEntityDraft, this.activeEntityLabel, this.pendingFocusField);
      this.pendingFocusField = undefined;
      return;
    }

    this.editorView.webview.html = renderEmptyEditorHtml();
  }

  private findEntityIndexByMap(kind: EditorEntityKind, target: YAML.YAMLMap): number | undefined {
    const collection = this.getCollectionByKind(kind, false);
    if (!collection || !Array.isArray(collection.items)) {
      return undefined;
    }

    const idx = collection.items.findIndex((item) => item === target);
    return idx >= 0 ? idx : undefined;
  }

  private getEntityId(kind: EditorEntityKind, entityMap: YAML.YAMLMap): string | undefined {
    if (kind === 'attacker') {
      const attackerId = entityMap.get('ID');
      if (typeof attackerId === 'string' && attackerId.length > 0) {
        return attackerId;
      }
      const name = entityMap.get('name');
      return typeof name === 'string' ? name : undefined;
    }

    const id = entityMap.get('ID');
    return typeof id === 'string' ? id : undefined;
  }

  private formatValue(value: unknown): string {
    if (typeof value === 'string') {
      return value;
    }

    if (value === null || value === undefined) {
      return 'null';
    }

    if (typeof value === 'object') {
      return JSON.stringify(value);
    }

    return String(value);
  }

  private getNodeStartOffset(node: unknown): number | undefined {
    if (!node || typeof node !== 'object') {
      return undefined;
    }

    const range = (node as { range?: unknown }).range;
    if (Array.isArray(range) && typeof range[0] === 'number') {
      return range[0];
    }

    return undefined;
  }

  private getNodeSelectionOffsets(node: unknown): [number, number] | null {
    if (!node || typeof node !== 'object') {
      return null;
    }

    const range = (node as { range?: unknown }).range;
    if (!Array.isArray(range) || typeof range[0] !== 'number') {
      return null;
    }

    const start = range[0];
    const candidateEnd = typeof range[1] === 'number' ? range[1] : start;
    const end = Math.max(start, candidateEnd);
    return [start, end];
  }

  private offsetToPosition(offset: number): vscode.Position {
    const source = this.tmSource ?? '';
    const safeOffset = Math.max(0, Math.min(offset, source.length));
    let line = 0;
    let character = 0;

    for (let i = 0; i < safeOffset; i++) {
      if (source[i] === '\n') {
        line += 1;
        character = 0;
      } else {
        character += 1;
      }
    }

    return new vscode.Position(line, character);
  }
}

export class YAMLTreeItem extends vscode.TreeItem {
  public modelNode?: ModelNode;
  public modelFilePath?: string;

  constructor(
    public readonly label: string,
    public readonly collapsibleState: vscode.TreeItemCollapsibleState,
    public children: YAMLTreeItem[] = [],
    public item?: YAML.YAMLSeq | YAML.YAMLMap | YAML.Scalar | unknown,
    public readonly nodeKind: 'root' | 'map' | 'seq' | 'scalar' | 'model' = 'scalar',
    public readonly category: TreeItemCategory = 'other'
  ) {
    super(label, collapsibleState);
    this.children = children;
    this.iconPath = nodeKind === 'model' ? undefined : new vscode.ThemeIcon(this.getIconId(nodeKind));
    if (nodeKind === 'root' && category !== 'other') {
      this.contextValue = `root-${category}`;
    }
  }

  private getIconId(kind: 'root' | 'map' | 'seq' | 'scalar' | 'model'): string {
    if (kind === 'model') {
      return 'repo';
    }

    if (kind === 'root') {
      return 'shield';
    }

    if (kind === 'map') {
      return 'symbol-object';
    }

    if (kind === 'seq') {
      return 'list-unordered';
    }

    return 'symbol-value';
  }
}