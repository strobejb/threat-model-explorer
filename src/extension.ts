// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';

import * as explr from './explorer';


// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	console.log('Congratulations, your extension "tmexp" is now active!');

	const customTreeDataProvider = new explr.YAMLTreeDataProvider();

	/**
	 * Register custom explorer 
	 */
	//
	const treeView = vscode.window.createTreeView('tmexpExplorer', {
		treeDataProvider: customTreeDataProvider
	});
	const threatEditorView = vscode.window.registerWebviewViewProvider('tmexpThreatEditorView', customTreeDataProvider);
	//vscode.commands.registerCommand('customExplorerView.refreshEntry', () => customTreeDataProvider.refresh());

	const refreshCommand = vscode.commands.registerCommand('tmexp.refresh', () => customTreeDataProvider.refresh());
	const newThreatCommand = vscode.commands.registerCommand('tmexp.newThreat', (item?: explr.YAMLTreeItem) => {
		customTreeDataProvider.createNewEntity('threat', item?.modelFilePath);
	});
	const newSecurityObjectiveCommand = vscode.commands.registerCommand('tmexp.newSecurityObjective', (item?: explr.YAMLTreeItem) => {
		customTreeDataProvider.createNewEntity('securityObjective', item?.modelFilePath);
	});
	const newAttackerCommand = vscode.commands.registerCommand('tmexp.newAttacker', (item?: explr.YAMLTreeItem) => {
		customTreeDataProvider.createNewEntity('attacker', item?.modelFilePath);
	});
	const collapseAllCommand = vscode.commands.registerCommand('tmexp.collapseAll', () => {
		void vscode.commands.executeCommand('workbench.actions.treeView.tmexpExplorer.collapseAll');
	});
	const revealCommand = vscode.commands.registerCommand('tmexp.revealNode', (item: explr.YAMLTreeItem) => {
		void customTreeDataProvider.revealTreeItem(item);
		customTreeDataProvider.openEntityEditor(item);
	});

	const selectionListener = treeView.onDidChangeSelection((event) => {
		if (customTreeDataProvider.isRevealSelection()) {
			return;
		}
		const selected = event.selection[0];
		if (!selected) {
			return;
		}

		void customTreeDataProvider.revealTreeItem(selected);
		customTreeDataProvider.openEntityEditor(selected);
	});

	customTreeDataProvider.setTreeView(treeView);

	// Track the active YAML file — switch the explorer when the user changes tabs
	const activeEditorListener = vscode.window.onDidChangeActiveTextEditor((editor) => {
		if (customTreeDataProvider.isEditorSwitchSuppressed()) {
			return;
		}
		if (!editor) {
			// Only clear when all text editors are truly closed, not on transient focus shifts
			if (vscode.window.visibleTextEditors.length === 0) {
				customTreeDataProvider.clear();
			}
			return;
		}
		const fsPath = editor.document.uri.fsPath;
		if (fsPath.endsWith('.yaml') || fsPath.endsWith('.yml')) {
			customTreeDataProvider.setActiveFile(fsPath);
		} else {
			customTreeDataProvider.clear();
		}
	});

	// Also sync to whatever is already open at activation time
	if (vscode.window.activeTextEditor) {
		const fsPath = vscode.window.activeTextEditor.document.uri.fsPath;
		if (fsPath.endsWith('.yaml') || fsPath.endsWith('.yml')) {
			customTreeDataProvider.setActiveFile(fsPath);
		}
	}

	const textChangeListener = vscode.workspace.onDidChangeTextDocument((event) => {
		if (event.document.uri.fsPath === customTreeDataProvider.getActiveFile()) {
			customTreeDataProvider.handleYamlSourceChanged();
		}
	});

	const saveListener = vscode.workspace.onDidSaveTextDocument((document) => {
		if (document.uri.fsPath === customTreeDataProvider.getActiveFile()) {
			customTreeDataProvider.handleYamlSourceChanged();
		}
	});

	const fileWatcher = vscode.workspace.createFileSystemWatcher('**/*.{yaml,yml}');
	const fileWatcherChangeListener = fileWatcher.onDidChange((uri) => {
		if (uri.fsPath === customTreeDataProvider.getActiveFile()) {
			customTreeDataProvider.handleYamlSourceChanged();
		}
	});
	const fileWatcherCreateListener = fileWatcher.onDidCreate((uri) => {
		if (uri.fsPath === customTreeDataProvider.getActiveFile()) {
			customTreeDataProvider.handleYamlSourceChanged();
		}
	});

	context.subscriptions.push(
		treeView,
		threatEditorView,
		refreshCommand,
		newThreatCommand,
		newSecurityObjectiveCommand,
		newAttackerCommand,
		collapseAllCommand,
		revealCommand,
		selectionListener,
		activeEditorListener,
		textChangeListener,
		saveListener,
		fileWatcher,
		fileWatcherChangeListener,
		fileWatcherCreateListener
	);
}

// This method is called when your extension is deactivated
export function deactivate() {}

// see: https://code.visualstudio.com/api/extension-guides/webview
//https://code.visualstudio.com/api/extension-guides/tree-view#tree-view-api-basics