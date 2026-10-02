# React Word Viewer

<div align="center">

**High-performance, pure client-side React component for viewing Microsoft Word documents (`.docx`, `.doc`) directly in the browser with zero server dependencies.**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-Ready-3178C6.svg)](https://www.typescriptlang.org/)
[![npm](https://img.shields.io/npm/v/react-word-viewer.svg)](https://www.npmjs.com/package/react-word-viewer)

[Live Docs](https://react-word-viewer-docs.vercel.app/) • [Playground](https://react-word-viewer-playground.vercel.app/)

</div>

---

## 🌟 Overview

`react-word-viewer` is a lightweight, zero-backend React library built for seamless Word document rendering. It decodes and renders `.docx` packages and `.doc` files purely on the client side using standard browser APIs, Web Streams, and CSS Grid/Flex layouts.

### Key Capabilities
- 🚀 **100% Client-Side**: No server conversion, no LibreOffice/Node daemon required.
- 🎨 **Theme Engine**: Light, Dark, and Sepia warm-reading modes.
- 🔍 **Instant Search**: Ctrl+F / Cmd+F document search with live highlights and match counts.
- 📑 **Interactive Table of Contents**: Automated outline generation from Word heading styles (H1–H6).
- 📖 **Dual Layout Modes**: Choose between Continuous Scrolling or Paginated views.
- 🌐 **Full RTL Support**: First-class support for Arabic, Hebrew, and bidirectional text rendering.
- 🖨️ **Print & Clipboard**: Built-in print formatting and one-click plain text copy.
- 📐 **Document Metadata**: Word count, character count, paragraph stats, and estimated reading time.

---

## 📦 Packages & Workspaces

This repository is organized as a monorepo:

| Directory | Name | Description |
|-----------|------|-------------|
| `packages/react-word-viewer` | [`react-word-viewer`](packages/react-word-viewer) | Core React Word viewer library |
| `apps/playground` | `playground` | Interactive testing playground for uploading and viewing documents |
| `apps/docs` | `docs` | Documentation and API reference portal |

---

## 🚀 Quick Start

### Installation

```bash
npm install react-word-viewer
```

### Basic Usage

```tsx
import React from 'react';
import { WordViewer } from 'react-word-viewer';

export default function App() {
  return (
    <div style={{ height: '100vh', width: '100%' }}>
      <WordViewer
        url="/sample.docx"
        title="Sample Document"
      />
    </div>
  );
}
```

---

## 🛠️ Development

### Prerequisites
- Node.js >= 18
- npm >= 9

### Setup

```bash
# Clone the repository
git clone https://github.com/hamzarihani/react-word-viewer.git
cd react-word-viewer

# Install all workspace dependencies
npm install

# Build all packages
npm run build:packages

# Run the playground locally
npm run dev:playground

# Run the documentation site locally
npm run dev:docs
```

---

## 📄 License

MIT © [Hamza Rihani](https://github.com/hamzarihani)
