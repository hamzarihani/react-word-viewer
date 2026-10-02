import { useState, type ChangeEvent, type DragEvent } from 'react';
import { WordViewer } from 'react-word-viewer';
import './App.css';

function App() {
  const [docUrl, setDocUrl] = useState<string | null>(null);
  const [docTitle, setDocTitle] = useState<string>('Document');
  const [locale, setLocale] = useState<'en' | 'ar'>('en');
  const [isDragging, setIsDragging] = useState(false);

  const handleFileUpload = (file: File) => {
    if (file) {
      const url = URL.createObjectURL(file);
      setDocUrl(url);
      setDocTitle(file.name);
    }
  };

  const handleInputChange = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFileUpload(file);
    }
  };

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      handleFileUpload(file);
    }
  };

  const loadSampleDocument = () => {
    setDocUrl('/sample.docx');
    setDocTitle('Sample Agreement & Feature Overview.docx');
  };

  return (
    <div className="playground-container">
      <header className="header">
        <div className="header-container">
          <div className="logo-container">
            <img src="/logo.svg" alt="React Word Viewer Logo" className="logo" />
            <div className="logo-text">
              <h1>react-word-viewer</h1>
              <span className="badge">Playground</span>
            </div>
          </div>
          <nav className="header-links">
            <a href="https://react-word-viewer-docs.vercel.app/" target="_blank" rel="noreferrer">
              Docs
            </a>
            <a href="https://github.com/hamzarihani/react-word-viewer" target="_blank" rel="noreferrer">
              GitHub
            </a>
            <a href="https://www.npmjs.com/package/react-word-viewer" target="_blank" rel="noreferrer">
              npm
            </a>
          </nav>
        </div>
      </header>

      <main className="playground-main">
        {!docUrl ? (
          <div className="upload-container">
            <div className="hero-text">
              <h2>Interactive Word Document Playground</h2>
              <p>
                Experience ultra-fast, client-side rendering for Microsoft Word (.docx and .doc) files.
                Zero server conversion, complete privacy, full rich-text styling.
              </p>
            </div>

            <div
              className={`dropzone ${isDragging ? 'dragging' : ''}`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <div className="dropzone-icon">
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" />
                  <line x1="12" y1="18" x2="12" y2="12" />
                  <line x1="9" y1="15" x2="15" y2="15" />
                </svg>
              </div>

              <h3>Drag & Drop your Word file here</h3>
              <p className="dropzone-hint">Supports .docx and legacy .doc files</p>

              <div className="action-buttons">
                <label className="btn btn-primary file-label">
                  Choose Word File
                  <input
                    type="file"
                    accept=".docx,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/msword"
                    onChange={handleInputChange}
                  />
                </label>

                <button type="button" onClick={loadSampleDocument} className="btn btn-secondary">
                  Load Sample Document
                </button>
              </div>
            </div>

            <div className="features-preview">
              <div className="feature-card">
                <span className="feature-icon">⚡</span>
                <h4>Zero Server Backend</h4>
                <p>Pure client-side unpacking and parsing. Files never leave the user browser.</p>
              </div>
              <div className="feature-card">
                <span className="feature-icon">🔍</span>
                <h4>Instant Full Search</h4>
                <p>Built-in search with text highlighting, occurrence counts, and keyboard shortcuts.</p>
              </div>
              <div className="feature-card">
                <span className="feature-icon">📑</span>
                <h4>Document Outline</h4>
                <p>Automatically discovers heading hierarchy and provides quick outline navigation.</p>
              </div>
              <div className="feature-card">
                <span className="feature-icon">🌐</span>
                <h4>Full RTL / Arabic Support</h4>
                <p>Bi-directional text rendering with specialized Arabic and English localizations.</p>
              </div>
            </div>
          </div>
        ) : (
          <div className="viewer-section">
            <div className="viewer-topbar">
              <div className="viewer-meta">
                <span className="doc-icon">📄</span>
                <span className="doc-title-text" title={docTitle}>
                  {docTitle}
                </span>
              </div>
              <div className="viewer-controls">
                <div className="lang-switcher">
                  <button
                    className={`lang-btn ${locale === 'en' ? 'active' : ''}`}
                    onClick={() => setLocale('en')}
                  >
                    English
                  </button>
                  <button
                    className={`lang-btn ${locale === 'ar' ? 'active' : ''}`}
                    onClick={() => setLocale('ar')}
                  >
                    العربية
                  </button>
                </div>
                <label className="btn btn-sm btn-secondary file-label-inline">
                  Change File
                  <input
                    type="file"
                    accept=".docx,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/msword"
                    onChange={handleInputChange}
                  />
                </label>
                <button onClick={() => setDocUrl(null)} className="btn btn-sm btn-close">
                  Close
                </button>
              </div>
            </div>
            <div className="viewer-container">
              <WordViewer
                url={docUrl}
                title={docTitle}
                locale={locale}
                onClose={() => setDocUrl(null)}
              />
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
