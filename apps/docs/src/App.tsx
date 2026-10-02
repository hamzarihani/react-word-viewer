import { useState } from 'react';
import './App.css';

function App() {
  const [copiedInstall, setCopiedInstall] = useState(false);
  const [copiedCode, setCopiedCode] = useState(false);
  const [pkgManager, setPkgManager] = useState<'npm' | 'yarn' | 'pnpm'>('npm');

  const getInstallCmd = () => {
    switch (pkgManager) {
      case 'yarn':
        return 'yarn add react-word-viewer';
      case 'pnpm':
        return 'pnpm add react-word-viewer';
      default:
        return 'npm i react-word-viewer';
    }
  };

  const handleCopyInstall = () => {
    navigator.clipboard.writeText(getInstallCmd());
    setCopiedInstall(true);
    setTimeout(() => setCopiedInstall(false), 2000);
  };

  const sampleCode = `import React from 'react';
import { WordViewer } from 'react-word-viewer';

export default function DocumentPreview() {
  return (
    <div style={{ height: '850px', width: '100%' }}>
      <WordViewer
        url="/documents/quarterly-report.docx"
        title="Quarterly Business Report"
        locale="en"
        onClose={() => console.log('Closed')}
      />
    </div>
  );
}`;

  const headlessCode = `import { parseWordDocument } from 'react-word-viewer';

async function extractDocData(fileUrl: string) {
  const res = await fetch(fileUrl);
  const buffer = await res.arrayBuffer();
  
  // Parse without rendering UI
  const doc = await parseWordDocument(buffer);
  
  console.log('Title:', doc.metadata.title);
  console.log('Word count:', doc.metadata.words);
  console.log('Outline headings:', doc.headings);
  console.log('Plain text:', doc.plainText);
}`;

  const handleCopyCode = () => {
    navigator.clipboard.writeText(sampleCode);
    setCopiedCode(true);
    setTimeout(() => setCopiedCode(false), 2000);
  };

  return (
    <div className="docs-page">
      <header className="header">
        <div className="header-container">
          <div className="logo-container">
            <img src="/logo.svg" alt="React Word Viewer Logo" className="logo" />
            <div className="logo-text">
              <h1>react-word-viewer</h1>
              <span className="badge">Docs</span>
            </div>
          </div>
          <nav className="header-links">
            <a href="#quickstart">Quick Start</a>
            <a href="#props">Props</a>
            <a href="#features">Features</a>
            <a href="#headless">Headless API</a>
            <a
              href="https://react-word-viewer-playground.vercel.app/"
              target="_blank"
              rel="noreferrer"
              className="playground-nav-btn"
            >
              Playground ↗
            </a>
            <a href="https://github.com/hamzarihani/react-word-viewer" target="_blank" rel="noreferrer">
              GitHub
            </a>
          </nav>
        </div>
      </header>

      <main className="docs-main">
        {/* HERO SECTION */}
        <section className="hero-section">
          <div className="hero-badges">
            <span className="pill">v1.0.0</span>
            <span className="pill">MIT License</span>
            <span className="pill">Zero Backend Required</span>
          </div>
          <h1 className="hero-title">
            The Pure Client-Side Word Document Viewer for React
          </h1>
          <p className="hero-description">
            Render, search, and navigate Microsoft Word (<code>.docx</code> and legacy <code>.doc</code>) files
            completely in the user’s browser. No node server, no cloud conversions, no privacy compromises.
          </p>
          <div className="hero-cta">
            <a href="#quickstart" className="btn btn-primary">
              Get Started
            </a>
            <a
              href="https://react-word-viewer-playground.vercel.app/"
              target="_blank"
              rel="noreferrer"
              className="btn btn-secondary"
            >
              Launch Live Playground
            </a>
          </div>
        </section>

        {/* QUICK START */}
        <section id="quickstart" className="docs-section">
          <h2 className="section-title">Quick Start</h2>
          <p className="section-subtitle">Install the package and import the component directly into your React app.</p>

          <div className="install-box">
            <div className="pkg-tabs">
              {(['npm', 'yarn', 'pnpm'] as const).map((m) => (
                <button
                  key={m}
                  className={`pkg-tab ${pkgManager === m ? 'active' : ''}`}
                  onClick={() => setPkgManager(m)}
                >
                  {m}
                </button>
              ))}
            </div>
            <div className="install-command">
              <code>{getInstallCmd()}</code>
              <button onClick={handleCopyInstall} className="copy-btn" aria-label="Copy install command">
                {copiedInstall ? 'Copied!' : 'Copy'}
              </button>
            </div>
          </div>

          <div className="code-card">
            <div className="code-header">
              <span>Basic Usage Example</span>
              <button onClick={handleCopyCode} className="copy-btn">
                {copiedCode ? 'Copied!' : 'Copy Code'}
              </button>
            </div>
            <pre className="code-content">
              <code>{sampleCode}</code>
            </pre>
          </div>
        </section>

        {/* PROPS */}
        <section id="props" className="docs-section">
          <h2 className="section-title">Component Props</h2>
          <p className="section-subtitle">
            Configure <code>WordViewer</code> to match your application theme, layout, and language preferences.
          </p>

          <div className="table-wrapper">
            <table className="props-table">
              <thead>
                <tr>
                  <th>Prop</th>
                  <th>Type</th>
                  <th>Default</th>
                  <th>Description</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td><code>url</code></td>
                  <td><code>string</code></td>
                  <td><span className="required">Required</span></td>
                  <td>URL, Object URL (<code>blob:</code>), or base64 data URI of the Word document.</td>
                </tr>
                <tr>
                  <td><code>title</code></td>
                  <td><code>string</code></td>
                  <td><code>'Word Document'</code></td>
                  <td>Document name displayed in header, print dialog, and info card.</td>
                </tr>
                <tr>
                  <td><code>locale</code></td>
                  <td><code>'en' | 'ar'</code></td>
                  <td><code>'en'</code></td>
                  <td>UI language. Switching to <code>'ar'</code> enables Arabic localization.</td>
                </tr>
                <tr>
                  <td><code>isRtl</code></td>
                  <td><code>boolean</code></td>
                  <td><code>auto</code></td>
                  <td>Force Right-to-Left or Left-to-Right orientation. Auto-detects Arabic text.</td>
                </tr>
                <tr>
                  <td><code>onClose</code></td>
                  <td><code>() =&gt; void</code></td>
                  <td><code>undefined</code></td>
                  <td>If provided, displays a close icon button in the viewer toolbar.</td>
                </tr>
                <tr>
                  <td><code>className</code></td>
                  <td><code>string</code></td>
                  <td><code>''</code></td>
                  <td>Custom CSS class applied to root container for styling adjustments.</td>
                </tr>
                <tr>
                  <td><code>dictionary</code></td>
                  <td><code>Partial&lt;WordViewerDictionary&gt;</code></td>
                  <td><code>undefined</code></td>
                  <td>Override any toolbar tooltip, label, or error message string.</td>
                </tr>
                <tr>
                  <td><code>transformUrl</code></td>
                  <td><code>(url: string) =&gt; string</code></td>
                  <td><code>undefined</code></td>
                  <td>Intercept and rewrite URLs before network fetching.</td>
                </tr>
              </tbody>
            </table>
          </div>
        </section>

        {/* FEATURES GRID */}
        <section id="features" className="docs-section">
          <h2 className="section-title">Built-in Capabilities</h2>
          <p className="section-subtitle">Everything you need for enterprise-grade document interaction.</p>

          <div className="features-grid">
            <div className="feature-item">
              <div className="feature-icon-badge">⚡</div>
              <h3>Pure Client-Side Engine</h3>
              <p>
                Zero backend server or LibreOffice processes required. Unpacks ZIP and parses WordprocessingML on
                the fly.
              </p>
            </div>
            <div className="feature-item">
              <div className="feature-icon-badge">🔍</div>
              <h3>Full-Text Search</h3>
              <p>
                Integrated search modal (Ctrl+F) with keyword highlighting, match counters, and instant keyboard
                navigation.
              </p>
            </div>
            <div className="feature-item">
              <div className="feature-icon-badge">📑</div>
              <h3>Document Outline</h3>
              <p>
                Automatically extracts headings (H1 through H6) into an interactive table of contents drawer.
              </p>
            </div>
            <div className="feature-item">
              <div className="feature-icon-badge">🌓</div>
              <h3>Multiple Themes</h3>
              <p>
                Switch between high-contrast Light mode, eye-friendly Sepia reading mode, and sleek Dark mode.
              </p>
            </div>
            <div className="feature-item">
              <div className="feature-icon-badge">🌐</div>
              <h3>First-Class RTL & Arabic</h3>
              <p>
                Auto-detects Arabic/Hebrew script, renders proper text alignment, and provides native Arabic UI.
              </p>
            </div>
            <div className="feature-item">
              <div className="feature-icon-badge">🖨️</div>
              <h3>Print & Clipboard</h3>
              <p>
                Format-optimized browser print preview and one-click copy of full extracted text to clipboard.
              </p>
            </div>
          </div>
        </section>

        {/* HEADLESS PARSER */}
        <section id="headless" className="docs-section">
          <h2 className="section-title">Headless Parser API</h2>
          <p className="section-subtitle">
            Need to extract document metadata, headings, or text in a background worker or custom UI? Use the headless
            parser.
          </p>

          <div className="code-card">
            <div className="code-header">
              <span>TypeScript Usage</span>
            </div>
            <pre className="code-content">
              <code>{headlessCode}</code>
            </pre>
          </div>
        </section>
      </main>

      <footer className="footer">
        <div className="footer-container">
          <div className="footer-left">
            <img src="/logo.svg" alt="React Word Viewer" className="footer-logo" />
            <p>React Word Viewer — Pure client-side document viewer.</p>
          </div>
          <div className="footer-links">
            <a href="https://github.com/hamzarihani/react-word-viewer" target="_blank" rel="noreferrer">
              GitHub
            </a>
            <a href="https://www.npmjs.com/package/react-word-viewer" target="_blank" rel="noreferrer">
              npm
            </a>
            <a href="https://react-word-viewer-playground.vercel.app/" target="_blank" rel="noreferrer">
              Playground
            </a>
          </div>
        </div>
      </footer>
    </div>
  );
}

export default App;
