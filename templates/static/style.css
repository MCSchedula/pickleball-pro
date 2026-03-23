:root {
    --primary: #059669;
    --primary-light: #10b981;
    --primary-dark: #047857;
    --secondary: #0284c7;
    --accent: #f59e0b;
    --danger: #dc2626;
    --success: #16a34a;
    
    --bg-dark: #0c1821;
    --bg-darker: #050a0f;
    --bg-card: #1a2332;
    --bg-elevated: #243447;
    
    --text-primary: #f8fafc;
    --text-secondary: #cbd5e1;
    --text-muted: #94a3b8;
    --border: #334155;
    --border-light: #475569;
}

* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    background: var(--bg-dark);
    color: var(--text-primary);
    line-height: 1.6;
}

.app-container {
    display: flex;
    min-height: 100vh;
}

/* Sidebar - Same as HTML version */
.sidebar {
    width: 280px;
    background: var(--bg-darker);
    border-right: 1px solid var(--border);
    display: flex;
    flex-direction: column;
    position: sticky;
    top: 0;
    height: 100vh;
    overflow-y: auto;
}

.logo {
    padding: 32px 24px;
    border-bottom: 1px solid var(--border);
}

.logo h1 {
    font-size: 24px;
    font-weight: 800;
    background: linear-gradient(135deg, var(--primary-light), var(--secondary));
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 4px;
}

.logo p {
    font-size: 13px;
    color: var(--text-muted);
}

.nav-menu {
    flex: 1;
    padding: 16px 12px;
}

.nav-section {
    margin-bottom: 24px;
}

.nav-section-title {
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: var(--text-muted);
    padding: 8px 12px;
    margin-bottom: 4px;
}

.nav-item {
    display: flex;
    align-items: center;
    padding: 12px 12px;
    margin: 2px 0;
    border-radius: 8px;
    color: var(--text-secondary);
    cursor: pointer;
    transition: all 0.2s;
    font-size: 14px;
    font-weight: 500;
    gap: 12px;
}

.nav-item:hover {
    background: var(--bg-card);
    color: var(--text-primary);
}

.nav-item.active {
    background: var(--primary);
    color: white;
    font-weight: 600;
}

.nav-item-icon {
    font-size: 18px;
    width: 20px;
    text-align: center;
}

.main-content {
    flex: 1;
    overflow-y: auto;
}

.top-bar {
    background: var(--bg-darker);
    border-bottom: 1px solid var(--border);
    padding: 20px 32px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    z-index: 10;
}

.top-bar-left h2 {
    font-size: 24px;
    font-weight: 700;
    margin-bottom: 4px;
}

.top-bar-left p {
    font-size: 14px;
    color: var(--text-muted);
}

.top-bar-actions {
    display: flex;
    gap: 12px;
}

.content-area {
    padding: 32px;
    max-width: 1800px;
}

.card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 24px;
}

.card-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 20px;
}

.card-title {
    font-size: 18px;
    font-weight: 700;
    color: var(--text-primary);
}

.card-subtitle {
    font-size: 14px;
    color: var(--text-muted);
    margin-top: 4px;
}

.btn {
    padding: 10px 20px;
    border: none;
    border-radius: 8px;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.2s;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    font-family: inherit;
}

.btn-primary {
    background: var(--primary);
    color: white;
}

.btn-primary:hover {
    background: var(--primary-dark);
    transform: translateY(-1px);
}

.btn-secondary {
    background: var(--bg-elevated);
    color: var(--text-primary);
    border: 1px solid var(--border);
}

.btn-secondary:hover {
    background: var(--bg-card);
}

.btn-sm {
    padding: 6px 12px;
    font-size: 13px;
}

.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
}

.stat-card {
    background: var(--bg-elevated);
    padding: 20px;
    border-radius: 12px;
    border: 1px solid var(--border);
}

.stat-value {
    font-size: 32px;
    font-weight: 800;
    color: var(--primary-light);
    margin-bottom: 4px;
}

.stat-label {
    font-size: 13px;
    color: var(--text-muted);
    font-weight: 500;
}

.hidden { display: none !important; }
.text-success { color: var(--success); }
.text-danger { color: var(--danger); }
.text-muted { color: var(--text-muted); }
.mb-24 { margin-bottom: 24px; }
.mt-16 { margin-top: 16px; }

.form-group {
    display: flex;
    flex-direction: column;
    gap: 8px;
    margin-bottom: 16px;
}

.form-label {
    font-size: 14px;
    font-weight: 600;
    color: var(--text-secondary);
}

.form-input, .form-select {
    padding: 12px 16px;
    background: var(--bg-dark);
    border: 1px solid var(--border);
    border-radius: 8px;
    color: var(--text-primary);
    font-size: 14px;
    font-family: inherit;
}

.form-input:focus, .form-select:focus {
    outline: none;
    border-color: var(--primary);
}

.form-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 20px;
    margin-bottom: 20px;
}

.player-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 12px;
}

.player-card {
    background: var(--bg-elevated);
    padding: 16px;
    border-radius: 10px;
    border: 2px solid var(--border);
    cursor: pointer;
    transition: all 0.2s;
    display: flex;
    align-items: center;
    gap: 12px;
}

.player-card:hover {
    border-color: var(--primary);
}

.player-card.selected {
    border-color: var(--primary);
    background: rgba(5, 150, 105, 0.1);
}

.player-checkbox {
    width: 20px;
    height: 20px;
    accent-color: var(--primary);
}

.player-info {
    flex: 1;
}

.player-name {
    font-weight: 600;
    font-size: 14px;
    margin-bottom: 4px;
}

.player-meta {
    font-size: 12px;
    color: var(--text-muted);
}

.level-badge {
    display: inline-block;
    padding: 2px 8px;
    background: var(--primary);
    color: white;
    border-radius: 4px;
    font-size: 12px;
    font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
}

.toast {
    position: fixed;
    bottom: 24px;
    right: 24px;
    padding: 16px 24px;
    background: var(--success);
    color: white;
    border-radius: 12px;
    font-weight: 600;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
    z-index: 1000;
    animation: slideUp 0.3s ease-out;
}

.toast.error {
    background: var(--danger);
}

@keyframes slideUp {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}
