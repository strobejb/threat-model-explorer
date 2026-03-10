# Threat Model Explorer

A Visual Studio Code extension for visualising and editing YAML-based threat models.

## Features

- **Explorer tree view** — Automatically discovers and displays hierarchical threat model structures from YAML files. Supports parent/child model relationships.
- **Inline entity editor** — Edit threats, security objectives, attackers, and model-level fields in a dedicated sidebar panel. Single-line text fields include an in-field commit button that appears when a change is made.
- **Live source synchronisation** — Edits in the panel are written back to the YAML document and reflected in the text editor. Switching editor tabs automatically updates the tree and panel.
- **CVSS v3.1 calculator** — Built-in base score calculator for threat entries, with a modal dialog for selecting metric values.
- **Draft workflow** — Create new threats, security objectives, and attackers via `[+]` buttons on root headings. Drafts are staged in the editor panel before being inserted into the YAML.
- **Collapse all** — Quickly collapse the entire explorer tree.

## Installation

1. Download the latest `tmexp.vsix` file from the [GitHub Releases](../../releases) page.
2. In VS Code, open the Command Palette (`Ctrl+Shift+P` / `Cmd+Shift+P`) and run **Extensions: Install from VSIX...**.
3. Select the downloaded `.vsix` file.
4. Reload VS Code when prompted.

## Getting Started

1. Install the extension.
2. Open a `.yaml` or `.yml` file that follows the threat model schema.
3. The **Threat Model Explorer** tree view appears in the sidebar. Click any entity to edit it in the panel below.

## Known Issues

See the issue tracker for open bugs and feature requests.

## Release Notes

See [CHANGELOG.md](CHANGELOG.md) for a full history of changes.
