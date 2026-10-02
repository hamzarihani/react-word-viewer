# react-word-viewer

A high-performance, browser-native React component for viewing, navigating, and interacting with Microsoft Word documents (`.docx` and legacy `.doc`) with zero server-side conversion or dependencies.

**Repository:** [github.com/hamzarihani/react-word-viewer](https://github.com/hamzarihani/react-word-viewer)

## Features

- ⚡ **Pure Client-Side**: Renders `.docx` documents completely in the browser using native web APIs. Zero server requirements.
- 📖 **Dual View Modes**: Switch between **Continuous Reading Mode** and **Paged Document Mode**.
- 🎨 **Theme System**: Instant switching between **Light**, **Sepia (Warm Reading)**, and **Dark** themes.
- 🔍 **Full-Text Search**: In-document search with match highlighting, occurrences count, and keyboard navigation (Ctrl+F / Cmd+F).
- 📑 **Interactive Table of Contents**: Automatically extracts document headings (H1–H6) into a navigatable outline drawer.
- 🖨️ **Print & Export**: Print preview and one-click full plain-text copy to clipboard.
- 🔍 **Smooth Zooming**: Scalable zoom controls (40% to 250%) with keyboard shortcuts.
- 🌐 **Full RTL & Arabic Support**: Native bidirectional Arabic/Hebrew text rendering, right-to-left layout orientation, and built-in Arabic/English localization.
- 📊 **Rich Elements**: Supports tables with backgrounds and cell spans, embedded images, colored text, highlights, lists, hyperlinks, and page breaks.
- 🛡️ **Zero Runtime Overhead**: Pre-bundled styles and types with TypeScript definitions included.

## Installation

```bash
npm install react-word-viewer
```

## Quick Start

```tsx
import React from 'react';
import { WordViewer } from 'react-word-viewer';

export default function DocumentPage() {
  return (
    <div style={{ height: '100vh', width: '100%' }}>
      <WordViewer
        url="https://example.com/agreement.docx"
        title="Agreement Document"
      />
    </div>
  );
}
```

## Props

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `url` | `string` | **Required** | URL or Blob URL of the Word (`.docx` / `.doc`) document. |
| `title` | `string` | `'Word Document'` | Title displayed in the toolbar, print title, and download filename. |
| `isRtl` | `boolean` | Auto-detected | Explicitly force RTL or LTR layout. |
| `locale` | `'en' \| 'ar'` | `'en'` | Default UI language (`'en'` for English, `'ar'` for Arabic). |
| `dictionary` | `Partial<WordViewerDictionary>` | `{}` | Custom localization dictionary overrides. |
| `className` | `string` | `''` | Optional CSS class name for the root container. |
| `onClose` | `() => void` | `undefined` | Optional close callback; displays an `X` button in the toolbar when provided. |
| `transformUrl` | `(url: string) => string` | `undefined` | Optional transformer for URL rewriting (e.g. proxying or cloud blobs). |

## Advanced Parsing Exports

You can also use the underlying document parser directly without the UI component:

```tsx
import { parseWordDocument, ZipReader } from 'react-word-viewer';

const response = await fetch('/sample.docx');
const buffer = await response.arrayBuffer();
const parsed = await parseWordDocument(buffer);

console.log('Document title:', parsed.metadata.title);
console.log('Word count:', parsed.metadata.words);
console.log('Headings:', parsed.headings);
```

## License

MIT © [Hamza Rihani](https://github.com/hamzarihani)
