<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Bank Transaction Processing System</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@200;400;500;700;900&family=Space+Grotesk:wght@400;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --primary: #008BB6;
    --primary-dark: #006d8e;
    --secondary: #00a2d4;
    --accent: #FCBA12;
    --accent-dark: #d9a00e;
    --bg: #F8F9FA;
    --bg-card: #ffffff;
    --text: #1A202C;
    --text-muted: #636E72;
    --border: #e2e8f0;
    --success: #27ae60;
    --danger: #e74c3c;
    --warning: #f39c12;
    --shadow: 0 4px 24px rgba(0,0,0,0.08);
    --shadow-lg: 0 12px 48px rgba(0,0,0,0.12);
    --radius: 12px;
    --radius-sm: 8px;
  }

  .dark {
    --primary: #00a2d4;
    --primary-dark: #008BB6;
    --secondary: #33b8e0;
    --bg: #1a1d23;
    --bg-card: #252830;
    --text: #e8ecf1;
    --text-muted: #8b95a5;
    --border: #363b45;
    --shadow: 0 4px 24px rgba(0,0,0,0.3);
    --shadow-lg: 0 12px 48px rgba(0,0,0,0.4);
  }

  * { margin: 0; padding: 0; box-sizing: border-box; }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    transition: background 0.4s, color 0.4s;
  }

  /* Animated background */
  .bg-pattern {
    position: fixed;
    inset: 0;
    z-index: 0;
    pointer-events: none;
    overflow: hidden;
  }
  .bg-pattern::before {
    content: '';
    position: absolute;
    width: 600px; height: 600px;
    background: radial-gradient(circle, rgba(0,139,182,0.08) 0%, transparent 70%);
    top: -200px; right: -100px;
    animation: float1 20s ease-in-out infinite;
  }
  .bg-pattern::after {
    content: '';
    position: absolute;
    width: 500px; height: 500px;
    background: radial-gradient(circle, rgba(252,186,18,0.06) 0%, transparent 70%);
    bottom: -150px; left: -100px;
    animation: float2 25s ease-in-out infinite;
  }
  @keyframes float1 {
    0%, 100% { transform: translate(0, 0) scale(1); }
    50% { transform: translate(-80px, 60px) scale(1.2); }
  }
  @keyframes float2 {
    0%, 100% { transform: translate(0, 0) scale(1); }
    50% { transform: translate(60px, -80px) scale(1.1); }
  }

  /* Header */
  .header {
    position: sticky;
    top: 0;
    z-index: 100;
    background: var(--primary);
    color: white;
    padding: 0 32px;
    height: 64px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    box-shadow: 0 2px 20px rgba(0,139,182,0.3);
  }
  .header-left {
    display: flex;
    align-items: center;
    gap: 16px;
  }
  .header-logo {
    width: 40px; height: 40px;
    background: var(--accent);
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 900;
    font-size: 18px;
    color: #1a1a1a;
    font-family: 'Space Grotesk', sans-serif;
  }
  .header-title {
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 20px;
    letter-spacing: -0.5px;
  }
  .header-subtitle {
    font-size: 12px;
    opacity: 0.8;
    font-weight: 400;
  }
  .header-right {
    display: flex;
    align-items: center;
    gap: 12px;
  }
  .icon-btn {
    width: 40px; height: 40px;
    border: none;
    border-radius: var(--radius-sm);
    background: rgba(255,255,255,0.15);
    color: white;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 16px;
    transition: all 0.2s;
  }
  .icon-btn:hover {
    background: rgba(255,255,255,0.25);
    transform: translateY(-1px);
  }

  /* Navigation tabs */
  .nav-tabs {
    display: flex;
    gap: 4px;
    padding: 16px 32px 0;
    position: relative;
    z-index: 10;
    overflow-x: auto;
    scrollbar-width: none;
  }
  .nav-tabs::-webkit-scrollbar { display: none; }

  .nav-tab {
    padding: 12px 20px;
    border: none;
    background: transparent;
    color: var(--text-muted);
    font-family: 'DM Sans', sans-serif;
    font-size: 14px;
    font-weight: 500;
    cursor: pointer;
    border-radius: var(--radius-sm) var(--radius-sm) 0 0;
    transition: all 0.25s;
    white-space: nowrap;
    display: flex;
    align-items: center;
    gap: 8px;
    position: relative;
  }
  .nav-tab:hover {
    color: var(--primary);
    background: rgba(0,139,182,0.06);
  }
  .nav-tab.active {
    color: var(--primary);
    background: var(--bg-card);
    font-weight: 700;
    box-shadow: 0 -2px 10px rgba(0,0,0,0.04);
  }
  .nav-tab.active::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
    background: var(--primary);
    border-radius: 3px 3px 0 0;
  }
  .nav-tab i { font-size: 15px; }

  /* Main content area */
  .main-content {
    position: relative;
    z-index: 10;
    padding: 24px 32px 60px;
  }

  .tab-panel {
    display: none;
    animation: fadeIn 0.35s ease;
  }
  .tab-panel.active { display: block; }

  @keyframes fadeIn {
    from { opacity: 0; transform: translateY(8px); }
    to { opacity: 1; transform: translateY(0); }
  }

  /* Cards */
  .card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    padding: 24px;
    margin-bottom: 20px;
    transition: background 0.4s, border-color 0.4s;
  }
  .card-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 20px;
    padding-bottom: 16px;
    border-bottom: 1px solid var(--border);
  }
  .card-title {
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 18px;
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .card-title i {
    color: var(--primary);
    font-size: 20px;
  }

  /* Upload zone */
  .upload-zone {
    border: 2px dashed var(--border);
    border-radius: var(--radius);
    padding: 60px 40px;
    text-align: center;
    cursor: pointer;
    transition: all 0.3s;
    position: relative;
    overflow: hidden;
  }
  .upload-zone::before {
    content: '';
    position: absolute;
    inset: 0;
    background: linear-gradient(135deg, rgba(0,139,182,0.03), rgba(252,186,18,0.03));
    opacity: 0;
    transition: opacity 0.3s;
  }
  .upload-zone:hover, .upload-zone.drag-over {
    border-color: var(--primary);
    background: rgba(0,139,182,0.04);
  }
  .upload-zone:hover::before, .upload-zone.drag-over::before { opacity: 1; }
  .upload-zone .icon {
    font-size: 48px;
    color: var(--primary);
    margin-bottom: 16px;
    position: relative;
    z-index: 1;
  }
  .upload-zone h3 {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 20px;
    margin-bottom: 8px;
    position: relative;
    z-index: 1;
  }
  .upload-zone p {
    color: var(--text-muted);
    font-size: 14px;
    position: relative;
    z-index: 1;
  }
  .upload-zone input[type="file"] {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
  }

  /* Buttons */
  .btn {
    padding: 10px 24px;
    border: none;
    border-radius: var(--radius-sm);
    font-family: 'DM Sans', sans-serif;
    font-size: 14px;
    font-weight: 600;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    transition: all 0.2s;
  }
  .btn:active { transform: scale(0.97); }
  .btn-primary {
    background: var(--primary);
    color: white;
  }
  .btn-primary:hover { background: var(--primary-dark); box-shadow: 0 4px 16px rgba(0,139,182,0.3); }
  .btn-accent {
    background: var(--accent);
    color: #1a1a1a;
  }
  .btn-accent:hover { background: var(--accent-dark); box-shadow: 0 4px 16px rgba(252,186,18,0.3); }
  .btn-success {
    background: var(--success);
    color: white;
  }
  .btn-success:hover { background: #219a52; }
  .btn-danger {
    background: var(--danger);
    color: white;
  }
  .btn-danger:hover { background: #c0392b; }
  .btn-outline {
    background: transparent;
    border: 1.5px solid var(--border);
    color: var(--text);
  }
  .btn-outline:hover { border-color: var(--primary); color: var(--primary); }
  .btn-sm { padding: 7px 14px; font-size: 13px; }
  .btn-lg { padding: 14px 32px; font-size: 16px; }
  .btn:disabled {
    opacity: 0.5;
    cursor: not-allowed;
    transform: none !important;
  }

  /* Progress bar */
  .progress-container {
    margin: 20px 0;
    display: none;
  }
  .progress-container.visible { display: block; }
  .progress-label {
    font-size: 13px;
    color: var(--text-muted);
    margin-bottom: 8px;
    display: flex;
    justify-content: space-between;
  }
  .progress-bar {
    height: 8px;
    background: var(--border);
    border-radius: 4px;
    overflow: hidden;
  }
  .progress-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--primary), var(--secondary));
    border-radius: 4px;
    width: 0%;
    transition: width 0.4s ease;
  }

  /* Status message */
  .status-msg {
    padding: 12px 16px;
    border-radius: var(--radius-sm);
    font-size: 14px;
    margin-top: 12px;
    display: none;
    align-items: center;
    gap: 10px;
  }
  .status-msg.visible { display: flex; }
  .status-msg.success { background: rgba(39,174,96,0.1); color: var(--success); border: 1px solid rgba(39,174,96,0.2); }
  .status-msg.error { background: rgba(231,76,60,0.1); color: var(--danger); border: 1px solid rgba(231,76,60,0.2); }
  .status-msg.info { background: rgba(0,139,182,0.1); color: var(--primary); border: 1px solid rgba(0,139,182,0.2); }
  .status-msg.warning { background: rgba(243,156,18,0.1); color: var(--warning); border: 1px solid rgba(243,156,18,0.2); }

  /* Data table */
  .table-wrapper {
    overflow-x: auto;
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
  }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
  }
  thead {
    background: var(--primary);
    color: white;
    position: sticky;
    top: 0;
    z-index: 5;
  }
  th {
    padding: 12px 14px;
    text-align: left;
    font-weight: 600;
    white-space: nowrap;
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  td {
    padding: 10px 14px;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
  }
  tbody tr {
    transition: background 0.15s;
  }
  tbody tr:hover {
    background: rgba(0,139,182,0.04);
  }
  .num { text-align: right; font-variant-numeric: tabular-nums; }
  .text-green { color: var(--success); }
  .text-red { color: var(--danger); }
  .text-accent { color: var(--accent); font-weight: 700; }
  .text-muted { color: var(--text-muted); }
  .text-primary { color: var(--primary); font-weight: 600; }

  /* Summary grid */
  .summary-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
  }
  .stat-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px;
    text-align: center;
    transition: all 0.25s;
    position: relative;
    overflow: hidden;
  }
  .stat-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
  }
  .stat-card.blue::before { background: var(--primary); }
  .stat-card.gold::before { background: var(--accent); }
  .stat-card.green::before { background: var(--success); }
  .stat-card.red::before { background: var(--danger); }
  .stat-card:hover {
    transform: translateY(-3px);
    box-shadow: var(--shadow-lg);
  }
  .stat-label {
    font-size: 12px;
    color: var(--text-muted);
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 8px;
  }
  .stat-value {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 28px;
    font-weight: 700;
    line-height: 1.1;
  }
  .stat-sub {
    font-size: 12px;
    color: var(--text-muted);
    margin-top: 4px;
  }

  /* Modal */
  .modal-overlay {
    position: fixed;
    inset: 0;
    background: rgba(0,0,0,0.5);
    backdrop-filter: blur(4px);
    z-index: 1000;
    display: none;
    align-items: center;
    justify-content: center;
    animation: fadeIn 0.2s;
  }
  .modal-overlay.visible {
    display: flex;
  }
  .modal {
    background: var(--bg-card);
    border-radius: var(--radius);
    box-shadow: var(--shadow-lg);
    width: 90%;
    max-width: 720px;
    max-height: 85vh;
    overflow-y: auto;
    animation: modalSlide 0.3s ease;
  }
  @keyframes modalSlide {
    from { opacity: 0; transform: translateY(20px) scale(0.97); }
    to { opacity: 1; transform: translateY(0) scale(1); }
  }
  .modal-header {
    padding: 20px 24px;
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: sticky;
    top: 0;
    background: var(--bg-card);
    z-index: 5;
    border-radius: var(--radius) var(--radius) 0 0;
  }
  .modal-header h3 {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 18px;
  }
  .modal-close {
    width: 36px; height: 36px;
    border: none;
    background: var(--border);
    border-radius: 50%;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 16px;
    color: var(--text-muted);
    transition: all 0.2s;
  }
  .modal-close:hover { background: var(--danger); color: white; }
  .modal-body { padding: 24px; }
  .modal-footer {
    padding: 16px 24px;
    border-top: 1px solid var(--border);
    display: flex;
    justify-content: flex-end;
    gap: 10px;
    position: sticky;
    bottom: 0;
    background: var(--bg-card);
    border-radius: 0 0 var(--radius) var(--radius);
  }

  /* Form elements */
  textarea.form-control {
    width: 100%;
    padding: 14px;
    border: 1.5px solid var(--border);
    border-radius: var(--radius-sm);
    font-family: 'Courier New', monospace;
    font-size: 13px;
    background: var(--bg);
    color: var(--text);
    resize: vertical;
    min-height: 180px;
    transition: border-color 0.2s;
  }
  textarea.form-control:focus {
    outline: none;
    border-color: var(--primary);
    box-shadow: 0 0 0 3px rgba(0,139,182,0.12);
  }

  /* Instruction panel */
  .instruction-panel {
    background: rgba(0,139,182,0.05);
    border: 1px solid rgba(0,139,182,0.15);
    border-radius: var(--radius-sm);
    padding: 16px 20px;
    margin-bottom: 20px;
    font-size: 14px;
    line-height: 1.7;
  }
  .instruction-panel strong { color: var(--primary); }
  .instruction-panel code {
    background: rgba(0,0,0,0.06);
    padding: 2px 6px;
    border-radius: 4px;
    font-size: 12px;
  }

  /* Tab sub-navigation */
  .sub-tabs {
    display: flex;
    gap: 4px;
    margin-bottom: 20px;
    background: var(--bg);
    padding: 4px;
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
    overflow-x: auto;
  }
  .sub-tab {
    padding: 8px 16px;
    border: none;
    background: transparent;
    color: var(--text-muted);
    font-size: 13px;
    font-weight: 500;
    cursor: pointer;
    border-radius: 6px;
    white-space: nowrap;
    transition: all 0.2s;
    font-family: 'DM Sans', sans-serif;
  }
  .sub-tab:hover { color: var(--text); background: rgba(0,0,0,0.04); }
  .sub-tab.active {
    background: var(--bg-card);
    color: var(--primary);
    font-weight: 700;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  }

  /* Toast notification */
  .toast-container {
    position: fixed;
    top: 80px;
    right: 24px;
    z-index: 2000;
    display: flex;
    flex-direction: column;
    gap: 10px;
  }
  .toast {
    padding: 14px 20px;
    border-radius: var(--radius-sm);
    font-size: 14px;
    font-weight: 500;
    box-shadow: var(--shadow-lg);
    display: flex;
    align-items: center;
    gap: 10px;
    animation: toastIn 0.3s ease;
    min-width: 300px;
  }
  .toast.success { background: var(--success); color: white; }
  .toast.error { background: var(--danger); color: white; }
  .toast.info { background: var(--primary); color: white; }
  @keyframes toastIn {
    from { opacity: 0; transform: translateX(40px); }
    to { opacity: 1; transform: translateX(0); }
  }

  /* Scrollable table container */
  .table-scroll {
    max-height: 500px;
    overflow: auto;
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
  }
  .table-scroll table { border: none; }
  .table-scroll thead { position: sticky; top: 0; z-index: 5; }

  /* Badge */
  .badge {
    display: inline-flex;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  .badge-blue { background: rgba(0,139,182,0.12); color: var(--primary); }
  .badge-gold { background: rgba(252,186,18,0.15); color: #b8860b; }
  .badge-green { background: rgba(39,174,96,0.12); color: var(--success); }

  /* Highlight row */
  .highlight-row { background: rgba(39,174,96,0.08) !important; }
  .highlight-row td { font-weight: 600; }

  /* Responsive */
  @media (max-width: 768px) {
    .header { padding: 0 16px; }
    .nav-tabs { padding: 12px 16px 0; }
    .main-content { padding: 16px; }
    .card { padding: 16px; }
    .summary-grid { grid-template-columns: repeat(2, 1fr); gap: 10px; }
    .stat-value { font-size: 22px; }
    .modal { width: 95%; }
    .upload-zone { padding: 40px 20px; }
  }
  @media (max-width: 480px) {
    .summary-grid { grid-template-columns: 1fr; }
    .header-subtitle { display: none; }
  }

  /* Reduced motion */
  @media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
      animation-duration: 0.01ms !important;
      transition-duration: 0.01ms !important;
    }
  }

  /* Scrollbar styling */
  ::-webkit-scrollbar { width: 8px; height: 8px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--text-muted); }

  /* Terminal detail row actions */
  .row-actions {
    display: flex;
    gap: 6px;
  }
  .row-actions button {
    width: 28px; height: 28px;
    border: 1px solid var(--border);
    background: var(--bg);
    border-radius: 6px;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 12px;
    color: var(--text-muted);
    transition: all 0.2s;
  }
  .row-actions button:hover {
    border-color: var(--primary);
    color: var(--primary);
  }

  /* File info */
  .file-info {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 14px 18px;
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    margin-top: 16px;
  }
  .file-info i { font-size: 28px; color: var(--primary); }
  .file-info .file-name { font-weight: 600; font-size: 14px; }
  .file-info .file-size { font-size: 12px; color: var(--text-muted); }

  /* Empty state */
  .empty-state {
    text-align: center;
    padding: 60px 20px;
    color: var(--text-muted);
  }
  .empty-state i { font-size: 48px; margin-bottom: 16px; opacity: 0.4; }
  .empty-state h3 { font-size: 18px; margin-bottom: 8px; color: var(--text); }
  .empty-state p { font-size: 14px; }
</style>
</head>
<body>

<div class="bg-pattern"></div>

<!-- Toast Container -->
<div class="toast-container" id="toastContainer"></div>

<!-- Header -->
<header class="header">
  <div class="header-left">
    <div class="header-logo">BT</div>
    <div>
      <div class="header-title">Bank Transaction Processor</div>
      <div class="header-subtitle">ATM &amp; ICBS Report System</div>
    </div>
  </div>
  <div class="header-right">
    <button class="icon-btn" id="darkModeBtn" title="Toggle Dark Mode" aria-label="Toggle dark mode">
      <i class="fas fa-moon"></i>
    </button>
    <button class="icon-btn" id="githubBtn" title="View on GitHub" aria-label="GitHub">
      <i class="fab fa-github"></i>
    </button>
  </div>
</header>

<!-- Navigation -->
<nav class="nav-tabs" role="tablist">
  <button class="nav-tab active" data-tab="atm" role="tab" aria-selected="true">
    <i class="fas fa-money-bill-wave"></i> ATM Daily
  </button>
  <button class="nav-tab" data-tab="daily" role="tab" aria-selected="false">
    <i class="fas fa-calendar-day"></i> Daily Report
  </button>
  <button class="nav-tab" data-tab="icbs" role="tab" aria-selected="false">
    <i class="fas fa-database"></i> ICBS Report
  </button>
  <button class="nav-tab" data-tab="converter" role="tab" aria-selected="false">
    <i class="fas fa-exchange-alt"></i> Converter
  </button>
</nav>

<!-- Main Content -->
<main class="main-content">

  <!-- ==================== ATM DAILY TAB ==================== -->
  <section class="tab-panel active" id="tab-atm" role="tabpanel">
    <div class="instruction-panel">
      <strong><i class="fas fa-info-circle"></i> Instructions:</strong>
      Upload an Excel file (.xls/.xlsx) containing ATM transaction data. The system will automatically:
      <br>1. Parse and categorize transactions (ATM Withdraw, POS, Ecom, Transfer)
      <br>2. Filter out reversal transactions by matching Ref No
      <br>3. Aggregate data per terminal ID
      <br>4. Prompt you to paste opening balances from Excel
      <br>5. Generate a downloadable report with Terminal Details and Summary
    </div>

    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="fas fa-file-upload"></i> Upload ATM Transaction File</div>
      </div>
      <div class="upload-zone" id="atmUploadZone">
        <input type="file" accept=".xls,.xlsx,.xlsm" id="atmFileInput" aria-label="Upload ATM transaction file">
        <div class="icon"><i class="fas fa-cloud-upload-alt"></i></div>
        <h3>Drop your Excel file here</h3>
        <p>Supports .xls, .xlsx, .xlsm formats</p>
      </div>
      <div class="file-info" id="atmFileInfo" style="display:none">
        <i class="fas fa-file-excel"></i>
        <div>
          <div class="file-name" id="atmFileName"></div>
          <div class="file-size" id="atmFileSize"></div>
        </div>
        <button class="btn btn-sm btn-outline" style="margin-left:auto" onclick="clearAtmFile()">Clear</button>
      </div>
      <div class="progress-container" id="atmProgress">
        <div class="progress-label">
          <span id="atmProgressText">Processing...</span>
          <span id="atmProgressPct">0%</span>
        </div>
        <div class="progress-bar"><div class="progress-fill" id="atmProgressFill"></div></div>
      </div>
      <div class="status-msg" id="atmStatus"></div>
      <div style="margin-top:16px; display:flex; gap:10px; flex-wrap:wrap">
        <button class="btn btn-primary btn-lg" id="atmProcessBtn" disabled onclick="processAtmData()">
          <i class="fas fa-cogs"></i> Process File
        </button>
      </div>
    </div>

    <!-- ATM Results -->
    <div id="atmResults" style="display:none">
      <div class="summary-grid" id="atmSummaryGrid"></div>

      <div class="card">
        <div class="card-header">
          <div class="card-title"><i class="fas fa-terminal"></i> Terminal Details</div>
          <button class="btn btn-sm btn-primary" onclick="exportAtmReport()">
            <i class="fas fa-download"></i> Export Report
          </button>
        </div>
        <div class="table-scroll">
          <table id="atmTerminalTable">
            <thead>
              <tr>
                <th>S/N</th><th>Terminal ID</th><th>Location</th>
                <th>Opening Bal.</th><th>Trans</th><th>Issuer</th>
                <th>Acquirer</th><th>Balance</th><th>Remark</th>
              </tr>
            </thead>
            <tbody id="atmTerminalBody"></tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <div class="card-title"><i class="fas fa-chart-bar"></i> Transaction Summary</div>
        </div>
        <div class="table-scroll">
          <table id="atmSummaryTable">
            <thead>
              <tr>
                <th>Sheet Name</th><th>Count</th><th>Total Amount</th><th>Total Fee</th>
                <th>Transaction Amt</th><th>Unique Cards</th><th>Total 1%</th>
                <th>1% / 2100</th><th>Amount / 2100</th>
              </tr>
            </thead>
            <tbody id="atmSummaryBody"></tbody>
          </table>
        </div>
      </div>

      <!-- Sub-tabs for transaction detail sheets -->
      <div class="card">
        <div class="card-header">
          <div class="card-title"><i class="fas fa-list-alt"></i> Transaction Details</div>
        </div>
        <div class="sub-tabs" id="atmSubTabs"></div>
        <div class="table-scroll" id="atmDetailTableContainer">
          <div class="empty-state">
            <i class="fas fa-hand-pointer"></i>
            <h3>Select a category</h3>
            <p>Click a tab above to view transaction details</p>
          </div>
        </div>
      </div>
    </div>
  </section>

  <!-- ==================== DAILY REPORT TAB ==================== -->
  <section class="tab-panel" id="tab-daily" role="tabpanel">
    <div class="instruction-panel">
      <strong><i class="fas fa-info-circle"></i> Daily Report:</strong>
      Upload an Excel file for daily transaction processing. Similar categorization as ATM Daily but without terminal-level aggregation.
    </div>

    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="fas fa-file-upload"></i> Upload Daily Report File</div>
      </div>
      <div class="upload-zone" id="dailyUploadZone">
        <input type="file" accept=".xls,.xlsx,.xlsm" id="dailyFileInput" aria-label="Upload daily report file">
        <div class="icon"><i class="fas fa-cloud-upload-alt"></i></div>
        <h3>Drop your daily Excel file here</h3>
        <p>Supports .xls, .xlsx, .xlsm formats</p>
      </div>
      <div class="file-info" id="dailyFileInfo" style="display:none">
        <i class="fas fa-file-excel"></i>
        <div>
          <div class="file-name" id="dailyFileName"></div>
          <div class="file-size" id="dailyFileSize"></div>
        </div>
        <button class="btn btn-sm btn-outline" style="margin-left:auto" onclick="clearDailyFile()">Clear</button>
      </div>
      <div class="progress-container" id="dailyProgress">
        <div class="progress-label">
          <span id="dailyProgressText">Processing...</span>
          <span id="dailyProgressPct">0%</span>
        </div>
        <div class="progress-bar"><div class="progress-fill" id="dailyProgressFill"></div></div>
      </div>
      <div class="status-msg" id="dailyStatus"></div>
      <div style="margin-top:16px">
        <button class="btn btn-primary btn-lg" id="dailyProcessBtn" disabled onclick="processDailyData()">
          <i class="fas fa-cogs"></i> Process File
        </button>
      </div>
    </div>

    <div id="dailyResults" style="display:none">
      <div class="card">
        <div class="card-header">
          <div class="card-title"><i class="fas fa-chart-bar"></i> Daily Summary</div>
          <button class="btn btn-sm btn-primary" onclick="exportDailyReport()">
            <i class="fas fa-download"></i> Export Report
          </button>
        </div>
        <div class="table-scroll">
          <table id="dailySummaryTable">
            <thead>
              <tr>
                <th>Sheet Name</th><th>Count</th><th>Total Amount</th><th>Total Fee</th>
                <th>Transaction Amt</th><th>Unique Cards</th><th>Total 1%</th>
                <th>1% / 2100</th><th>Amount / 2100</th>
              </tr>
            </thead>
            <tbody id="dailySummaryBody"></tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <div class="card-title"><i class="fas fa-list-alt"></i> Transaction Details</div>
        </div>
        <div class="sub-tabs" id="dailySubTabs"></div>
        <div class="table-scroll" id="dailyDetailTableContainer">
          <div class="empty-state">
            <i class="fas fa-hand-pointer"></i>
            <h3>Select a category</h3>
            <p>Click a tab above to view transaction details</p>
          </div>
        </div>
      </div>
    </div>
  </section>

  <!-- ==================== ICBS REPORT TAB ==================== -->
  <section class="tab-panel" id="tab-icbs" role="tabpanel">
    <div class="instruction-panel">
      <strong><i class="fas fa-info-circle"></i> ICBS Report:</strong>
      Upload ICBS transaction data. The system will parse, validate, and allow you to generate structured reports with filtering and export capabilities.
    </div>

    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="fas fa-file-upload"></i> Upload ICBS Data</div>
      </div>
      <div class="upload-zone" id="icbsUploadZone">
        <input type="file" accept=".xls,.xlsx,.xlsm,.csv" id="icbsFileInput" aria-label="Upload ICBS data file">
        <div class="icon"><i class="fas fa-cloud-upload-alt"></i></div>
        <h3>Drop your ICBS file here</h3>
        <p>Supports .xls, .xlsx, .xlsm, .csv formats</p>
      </div>
      <div class="file-info" id="icbsFileInfo" style="display:none">
        <i class="fas fa-file-excel"></i>
        <div>
          <div class="file-name" id="icbsFileName"></div>
          <div class="file-size" id="icbsFileSize"></div>
        </div>
        <button class="btn btn-sm btn-outline" style="margin-left:auto" onclick="clearIcbsFile()">Clear</button>
      </div>
      <div class="status-msg" id="icbsStatus"></div>
      <div style="margin-top:16px; display:flex; gap:10px; flex-wrap:wrap">
        <button class="btn btn-primary btn-lg" id="icbsProcessBtn" disabled onclick="processIcbsData()">
          <i class="fas fa-cogs"></i> Process ICBS Data
        </button>
        <button class="btn btn-accent btn-lg" id="icbsExportBtn" disabled onclick="exportIcbsReport()">
          <i class="fas fa-download"></i> Export
        </button>
      </div>
    </div>

    <!-- ICBS Filter bar -->
    <div class="card" id="icbsFilterCard" style="display:none">
      <div class="card-header">
        <div class="card-title"><i class="fas fa-filter"></i> Filters</div>
        <button class="btn btn-sm btn-outline" onclick="clearIcbsFilters()">Reset Filters</button>
      </div>
      <div style="display:flex; gap:12px; flex-wrap:wrap; align-items:end">
        <div style="flex:1; min-width:180px">
          <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px">Account Number</label>
          <input type="text" id="icbsFilterAccount" placeholder="Search account..." style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:13px" oninput="applyIcbsFilters()">
        </div>
        <div style="flex:1; min-width:180px">
          <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px">Description</label>
          <input type="text" id="icbsFilterDesc" placeholder="Search description..." style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:13px" oninput="applyIcbsFilters()">
        </div>
        <div style="flex:1; min-width:120px">
          <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px">Type</label>
          <select id="icbsFilterType" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:13px" onchange="applyIcbsFilters()">
            <option value="">All</option>
            <option value="D">Debit</option>
            <option value="C">Credit</option>
          </select>
        </div>
        <div style="flex:1; min-width:140px">
          <label style="font-size:12px; color:var(--text-muted); display:block; margin-bottom:4px">Terminal ID</label>
          <input type="text" id="icbsFilterTerminal" placeholder="Search TID..." style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:6px;background:var(--bg);color:var(--text);font-size:13px" oninput="applyIcbsFilters()">
        </div>
      </div>
    </div>

    <!-- ICBS Summary -->
    <div class="summary-grid" id="icbsSummaryGrid" style="display:none"></div>

    <!-- ICBS Table -->
    <div class="card" id="icbsTableCard" style="display:none">
      <div class="card-header">
        <div class="card-title"><i class="fas fa-table"></i> ICBS Transactions</div>
        <span class="badge badge-blue" id="icbsRowCount">0 rows</span>
      </div>
      <div class="table-scroll" style="max-height:600px">
        <table id="icbsTable">
          <thead id="icbsTableHead"></thead>
          <tbody id="icbsTableBody"></tbody>
        </table>
      </div>
    </div>
  </section>

  <!-- ==================== CONVERTER TAB ==================== -->
  <section class="tab-panel" id="tab-converter" role="tabpanel">
    <div class="instruction-panel">
      <strong><i class="fas fa-info-circle"></i> XLS to XLSX Converter:</strong>
      Upload old-format .xls files to convert them to modern .xlsx format. All sheets and data will be preserved.
    </div>

    <div class="card">
      <div class="card-header">
        <div class="card-title"><i class="fas fa-exchange-alt"></i> Convert .xls to .xlsx</div>
      </div>
      <div class="upload-zone" id="converterUploadZone">
        <input type="file" accept=".xls" id="converterFileInput" aria-label="Upload XLS file to convert">
        <div class="icon"><i class="fas fa-file-import"></i></div>
        <h3>Drop your .xls file here</h3>
        <p>Only .xls format is accepted for conversion</p>
      </div>
      <div class="file-info" id="converterFileInfo" style="display:none">
        <i class="fas fa-file-excel"></i>
        <div>
          <div class="file-name" id="converterFileName"></div>
          <div class="file-size" id="converterFileSize"></div>
        </div>
      </div>
      <div class="status-msg" id="converterStatus"></div>
      <div style="margin-top:16px">
        <button class="btn btn-accent btn-lg" id="converterBtn" disabled onclick="convertFile()">
          <i class="fas fa-sync-alt"></i> Convert & Download
        </button>
      </div>
    </div>
  </section>

</main>

<!-- Opening Balances Modal -->
<div class="modal-overlay" id="openingBalanceModal">
  <div class="modal">
    <div class="modal-header">
      <h3><i class="fas fa-coins" style="color:var(--accent);margin-right:8px"></i>Paste Opening Balances</h3>
      <button class="modal-close" onclick="closeOpeningBalanceModal()" aria-label="Close modal"><i class="fas fa-times"></i></button>
    </div>
    <div class="modal-body">
      <p style="margin-bottom:12px; color:var(--text-muted); font-size:14px; line-height:1.6">
        Copy two columns from Excel and paste below:<br>
        <strong>Column 1:</strong> Terminal ID &nbsp;|&nbsp; <strong>Column 2:</strong> Opening Balance
      </p>
      <div style="background:var(--bg); padding:12px 16px; border-radius:6px; margin-bottom:16px; font-family:'Courier New',monospace; font-size:12px; color:var(--text-muted); line-height:1.8">
        Example format:<br>
        09950001&nbsp;&nbsp;&nbsp;&nbsp;150000.00<br>
        09950002&nbsp;&nbsp;&nbsp;&nbsp;200000.00<br>
        09950003&nbsp;&nbsp;&nbsp;&nbsp;175000.00
      </div>
      <textarea class="form-control" id="openingBalanceText" placeholder="Paste your data here..." rows="10"></textarea>
      <div class="status-msg" id="openingBalanceStatus" style="margin-top:12px"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-outline" onclick="closeOpeningBalanceModal()">Cancel</button>
      <button class="btn btn-success" onclick="submitOpeningBalances()">
        <i class="fas fa-check"></i> Parse & Submit
      </button>
    </div>
  </div>
</div>

<script>
// ===================== GLOBAL STATE =====================
const TERMINAL_IDS = [
  "09950001","09950002","09950003","09950004","09950005","09950006","09950007","09950008","09950009","09950010",
  "09950013","09950014","09950015","09950016","09950017","09950018","09950019","09950021","09950024","09950025",
  "09950026","09950027","09950028","09950029"
];

const TERMINAL_LOCATIONS = {
  "09950001":"Dagon","09950002":"Botahtaung","09950003":"MIP","09950004":"MDY","09950005":"NPT MOC 40",
  "09950006":"Pathein","09950007":"Naypyitaw Br","09950008":"Mawlamyaing","09950009":"Shwe Pauk Kan","09950010":"Magway",
  "09950013":"Dagon (2)","09950014":"Monywa","09950015":"Hpa-an","09950016":"Dawei","09950017":"DSK",
  "09950018":"BGO","09950019":"MKN","09950021":"HTY","09950024":"TGI","09950025":"MICT Park",
  "09950026":"NPT MOC 11","09950027":"NPT MOC 11 (2)","09950028":"NPT MOC 40 (2)","09950029":"NPT Hluttaw"
};

const TRANSACTION_CATEGORIES = [
  "ATM_Card_950316","ATM_Issuer_Withdraw","POS_Issuer_Sale",
  "Ecom_Issuer_Sale","ATM_Withdraw_Fee_NonZero","ATM_Transfer_CardToAccount"
];

// Store processed data for export
let atmProcessedData = null;
let dailyProcessedData = null;
let icbsProcessedData = null;
let atmFileBuffer = null;
let dailyFileBuffer = null;
let icbsFileBuffer = null;
let converterFileBuffer = null;

// ===================== UTILITY FUNCTIONS =====================

function toFloatSafe(value) {
  if (value === null || value === undefined || value === '') return 0.0;
  try {
    const s = String(value).replace(/,/g, '').replace(/MMK/gi, '').trim();
    return parseFloat(s) || 0.0;
  } catch { return 0.0; }
}

function formatNumber(n, decimals = 2) {
  return Number(n).toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function findColumnIndex(headers, keyword) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && String(headers[i]).toLowerCase().includes(keyword.toLowerCase())) return i;
  }
  return -1;
}

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  const icons = { success: 'fa-check-circle', error: 'fa-exclamation-circle', info: 'fa-info-circle' };
  toast.innerHTML = `<i class="fas ${icons[type] || icons.info}"></i> ${message}`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(40px)';
    toast.style.transition = 'all 0.3s';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

function showStatus(elId, msg, type) {
  const el = document.getElementById(elId);
  el.className = `status-msg visible ${type}`;
  const icons = { success: 'fa-check-circle', error: 'fa-times-circle', info: 'fa-info-circle', warning: 'fa-exclamation-triangle' };
  el.innerHTML = `<i class="fas ${icons[type] || icons.info}"></i> ${msg}`;
}

function hideStatus(elId) {
  const el = document.getElementById(elId);
  el.className = 'status-msg';
}

function setProgress(prefix, value, text) {
  const fill = document.getElementById(`${prefix}ProgressFill`);
  const pct = document.getElementById(`${prefix}ProgressPct`);
  const txt = document.getElementById(`${prefix}ProgressText`);
  const container = document.getElementById(`${prefix}Progress`);
  container.classList.add('visible');
  fill.style.width = value + '%';
  pct.textContent = Math.round(value) + '%';
  if (text) txt.textContent = text;
}

function hideProgress(prefix) {
  document.getElementById(`${prefix}Progress`).classList.remove('visible');
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

// Parse Excel file using SheetJS
function parseExcelFile(buffer) {
  return new Promise((resolve, reject) => {
    try {
      const data = new Uint8Array(buffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheet];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      resolve({ workbook, jsonData, sheetName: firstSheet });
    } catch (e) {
      reject(e);
    }
  });
}

// Categorize a single transaction row
function categorizeTransaction(rowText) {
  if (rowText.includes('ATM Withdraw') && rowText.includes('950316')) return 'ATM_Card_950316';
  if (rowText.includes('ATM Issuer Withdraw')) return 'ATM_Issuer_Withdraw';
  if (rowText.includes('POS Issuer Sale')) return 'POS_Issuer_Sale';
  if (rowText.includes('Ecom Issuer Sale')) return 'Ecom_Issuer_Sale';
  if (rowText.includes('ATM Withdraw') && !rowText.includes('950316')) return 'ATM_Withdraw_Fee_NonZero';
  if (rowText.includes('ATM Transfer') && (rowText.includes('CardToAccount') || rowText.includes('Card to Account'))) return 'ATM_Transfer_CardToAccount';
  return null;
}

// Process rows into categorized sheets
function processRows(rows, header, startRow, amountIdx, feeIdx, cardIdx, refnoIdx) {
  // Collect reversal ref nos first
  const reversalRefNos = new Set();
  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    const txt = row.map(c => String(c || '').trim()).join(' ');
    const refNo = String(row[refnoIdx] || '').trim();
    if (txt.toLowerCase().includes('reversal') && refNo) {
      reversalRefNos.add(refNo);
    }
  }

  const sheets = {};
  TRANSACTION_CATEGORIES.forEach(c => sheets[c] = []);

  for (let i = startRow; i < rows.length; i++) {
    const row = rows[i];
    const txt = row.map(c => String(c || '').trim()).join(' ');
    const refNo = String(row[refnoIdx] || '').trim();

    // Skip reversals
    if (refNo && reversalRefNos.has(refNo)) continue;

    const category = categorizeTransaction(txt);
    if (category) {
      sheets[category].push({ rowIndex: i, data: row });
    }
  }

  return sheets;
}

// Calculate summary for a sheet
function calculateSheetSummary(rows, amountIdx, feeIdx, cardIdx, sheetName) {
  let totalAmount = 0, totalFee = 0, count = 0;
  const uniqueCards = new Set();

  rows.forEach(r => {
    count++;
    totalAmount += toFloatSafe(r.data[amountIdx]);
    totalFee += toFloatSafe(r.data[feeIdx]);
    if (cardIdx >= 0 && r.data[cardIdx]) {
      uniqueCards.add(String(r.data[cardIdx]).trim());
    }
  });

  const transactionAmount = totalAmount - totalFee;
  const totalUseCard = uniqueCards.size;
  let total1Percent;
  if (sheetName === 'ATM_Issuer_Withdraw') {
    total1Percent = (transactionAmount * 0.01) + transactionAmount;
  } else {
    total1Percent = transactionAmount * 0.01;
  }
  const total1Div2100 = total1Percent / 2100;
  const totalDiv2100 = totalAmount / 2100;

  return {
    sheetName, count, totalAmount, totalFee, transactionAmount,
    totalUseCard, total1Percent, total1Div2100, totalDiv2100
  };
}

// ===================== DARK MODE =====================
const darkModeBtn = document.getElementById('darkModeBtn');
let isDark = false;

darkModeBtn.addEventListener('click', () => {
  isDark = !isDark;
  document.body.classList.toggle('dark', isDark);
  darkModeBtn.innerHTML = isDark ? '<i class="fas fa-sun"></i>' : '<i class="fas fa-moon"></i>';
});

// GitHub button
document.getElementById('githubBtn').addEventListener('click', () => {
  window.open('https://github.com', '_blank');
});

// ===================== TAB NAVIGATION =====================
document.querySelectorAll('.nav-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.nav-tab').forEach(t => {
      t.classList.remove('active');
      t.setAttribute('aria-selected', 'false');
    });
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    tab.setAttribute('aria-selected', 'true');
    document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
  });
});

// ===================== FILE UPLOAD HANDLING =====================

// Drag & drop setup for all upload zones
function setupUploadZone(zoneId, inputId, onFileLoaded) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);

  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) onFileLoaded(files[0]);
  });
  input.addEventListener('change', (e) => {
    if (e.target.files.length > 0) onFileLoaded(e.target.files[0]);
  });
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// --- ATM File Upload ---
setupUploadZone('atmUploadZone', 'atmFileInput', async (file) => {
  try {
    atmFileBuffer = await readFileAsArrayBuffer(file);
    document.getElementById('atmFileInfo').style.display = 'flex';
    document.getElementById('atmFileName').textContent = file.name;
    document.getElementById('atmFileSize').textContent = formatFileSize(file.size);
    document.getElementById('atmProcessBtn').disabled = false;
    hideStatus('atmStatus');
    showToast(`File "${file.name}" loaded successfully`, 'success');
  } catch (e) {
    showStatus('atmStatus', 'Failed to read file: ' + e.message, 'error');
  }
});

function clearAtmFile() {
  atmFileBuffer = null;
  document.getElementById('atmFileInput').value = '';
  document.getElementById('atmFileInfo').style.display = 'none';
  document.getElementById('atmProcessBtn').disabled = true;
  document.getElementById('atmResults').style.display = 'none';
  hideStatus('atmStatus');
  hideProgress('atm');
}

// --- Daily File Upload ---
setupUploadZone('dailyUploadZone', 'dailyFileInput', async (file) => {
  try {
    dailyFileBuffer = await readFileAsArrayBuffer(file);
    document.getElementById('dailyFileInfo').style.display = 'flex';
    document.getElementById('dailyFileName').textContent = file.name;
    document.getElementById('dailyFileSize').textContent = formatFileSize(file.size);
    document.getElementById('dailyProcessBtn').disabled = false;
    hideStatus('dailyStatus');
    showToast(`File "${file.name}" loaded successfully`, 'success');
  } catch (e) {
    showStatus('dailyStatus', 'Failed to read file: ' + e.message, 'error');
  }
});

function clearDailyFile() {
  dailyFileBuffer = null;
  document.getElementById('dailyFileInput').value = '';
  document.getElementById('dailyFileInfo').style.display = 'none';
  document.getElementById('dailyProcessBtn').disabled = true;
  document.getElementById('dailyResults').style.display = 'none';
  hideStatus('dailyStatus');
  hideProgress('daily');
}

// --- ICBS File Upload ---
setupUploadZone('icbsUploadZone', 'icbsFileInput', async (file) => {
  try {
    icbsFileBuffer = await readFileAsArrayBuffer(file);
    document.getElementById('icbsFileInfo').style.display = 'flex';
    document.getElementById('icbsFileName').textContent = file.name;
    document.getElementById('icbsFileSize').textContent = formatFileSize(file.size);
    document.getElementById('icbsProcessBtn').disabled = false;
    hideStatus('icbsStatus');
    showToast(`File "${file.name}" loaded successfully`, 'success');
  } catch (e) {
    showStatus('icbsStatus', 'Failed to read file: ' + e.message, 'error');
  }
});

function clearIcbsFile() {
  icbsFileBuffer = null;
  icbsProcessedData = null;
  document.getElementById('icbsFileInput').value = '';
  document.getElementById('icbsFileInfo').style.display = 'none';
  document.getElementById('icbsProcessBtn').disabled = true;
  document.getElementById('icbsExportBtn').disabled = true;
  document.getElementById('icbsFilterCard').style.display = 'none';
  document.getElementById('icbsSummaryGrid').style.display = 'none';
  document.getElementById('icbsTableCard').style.display = 'none';
  hideStatus('icbsStatus');
}

// --- Converter File Upload ---
setupUploadZone('converterUploadZone', 'converterFileInput', async (file) => {
  if (!file.name.toLowerCase().endsWith('.xls')) {
    showStatus('converterStatus', 'Please select a .xls file (old Excel format)', 'error');
    return;
  }
  try {
    converterFileBuffer = await readFileAsArrayBuffer(file);
    document.getElementById('converterFileInfo').style.display = 'flex';
    document.getElementById('converterFileName').textContent = file.name;
    document.getElementById('converterFileSize').textContent = formatFileSize(file.size);
    document.getElementById('converterBtn').disabled = false;
    hideStatus('converterStatus');
    showToast(`File "${file.name}" loaded for conversion`, 'success');
  } catch (e) {
    showStatus('converterStatus', 'Failed to read file: ' + e.message, 'error');
  }
});

// ===================== ATM PROCESSING =====================
async function processAtmData() {
  if (!atmFileBuffer) return;

  setProgress('atm', 10, 'Parsing Excel file...');

  let parsed;
  try {
    parsed = await parseExcelFile(atmFileBuffer);
  } catch (e) {
    showStatus('atmStatus', 'Failed to parse Excel file: ' + e.message, 'error');
    hideProgress('atm');
    return;
  }

  const rows = parsed.jsonData;
  if (rows.length < 2) {
    showStatus('atmStatus', 'File appears to be empty or has no data rows.', 'error');
    hideProgress('atm');
    return;
  }

  setProgress('atm', 30, 'Finding header columns...');

  // Find header row
  let header = null;
  let headerRow = 0;
  for (let r = 0; r < Math.min(3, rows.length); r++) {
    const testHeader = rows[r];
    const headerText = testHeader.map(h => String(h || '').toLowerCase()).join(' ');
    if (headerText.includes('terminal') || headerText.includes('amount') || headerText.includes('ref')) {
      header = testHeader;
      headerRow = r;
      break;
    }
  }
  if (!header) header = rows[0];

  const amountIdx = findColumnIndex(header, 'amount');
  const feeIdx = findColumnIndex(header, 'fee');
  const cardIdx = findColumnIndex(header, 'card');
  const refnoIdx = findColumnIndex(header, 'ref');
  let terminalIdIdx = -1;
  for (const kw of ['terminal id', 'terminal', 'tid']) {
    terminalIdIdx = findColumnIndex(header, kw);
    if (terminalIdIdx >= 0) break;
  }

  if (amountIdx < 0 || feeIdx < 0 || refnoIdx < 0) {
    showStatus('atmStatus', `Could not find required columns (amount, fee, ref no). Detected headers: ${header.map((h,i) => `${i+1}.${h}`).join(', ')}`, 'error');
    hideProgress('atm');
    return;
  }
  if (terminalIdIdx < 0) {
    showStatus('atmStatus', `Could not find 'terminal id' column. Detected headers: ${header.map((h,i) => `${i+1}.${h}`).join(', ')}`, 'error');
    hideProgress('atm');
    return;
  }

  setProgress('atm', 50, 'Categorizing transactions...');

  const sheets = processRows(rows, header, headerRow + 1, amountIdx, feeIdx, cardIdx, refnoIdx);

  // Terminal aggregation
  const terminalAggregation = {};
  const terminalTransactions = {};
  TERMINAL_IDS.forEach(tid => {
    terminalAggregation[tid] = {
      atm_issuer_withdraw_count: 0, atm_issuer_withdraw_amount: 0,
      atm_card_950316_amount: 0, atm_withdraw_fee_nonzero_amount: 0,
      atm_withdraw_fee_nonzero_trx_amount: 0, atm_withdraw_nonzero_count: 0
    };
    terminalTransactions[tid] = [];
  });

  let posEcomCount = 0, posEcomAmount = 0;
  let issuerCommissionCount = 0, issuerCommission = 0;
  let cardToCardCount = 0, cardToCardAmount = 0;

  // Re-scan for terminal-level aggregation
  const reversalRefNos = new Set();
  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i];
    const txt = row.map(c => String(c || '').trim()).join(' ');
    const refNo = String(row[refnoIdx] || '').trim();
    if (txt.toLowerCase().includes('reversal') && refNo) reversalRefNos.add(refNo);
  }

  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i];
    const txt = row.map(c => String(c || '').trim()).join(' ');
    const refNo = String(row[refnoIdx] || '').trim();
    if (refNo && reversalRefNos.has(refNo)) continue;

    const tid = String(row[terminalIdIdx] || '').trim();
    const amount = toFloatSafe(row[amountIdx]);
    const fee = toFloatSafe(row[feeIdx]);

    if (TERMINAL_IDS.includes(tid)) {
      terminalTransactions[tid].push(row);

      if (txt.includes('ATM Withdraw') && txt.includes('950316')) {
        terminalAggregation[tid].atm_card_950316_amount += amount;
      } else if (txt.includes('ATM Issuer Withdraw')) {
        if (amount > 0) terminalAggregation[tid].atm_issuer_withdraw_count++;
        terminalAggregation[tid].atm_issuer_withdraw_amount += amount;
        if (amount > 0) {
          issuerCommissionCount++;
          issuerCommission += amount * 0.01;
        }
      } else if (txt.includes('POS Issuer Sale') || txt.includes('Ecom Issuer Sale')) {
        if (amount > 0) {
          posEcomCount++;
          posEcomAmount += amount;
        }
      } else if (txt.includes('ATM Withdraw') && !txt.includes('950316')) {
        terminalAggregation[tid].atm_withdraw_fee_nonzero_trx_amount += amount;
        terminalAggregation[tid].atm_withdraw_fee_nonzero_amount += fee;
      }

      if (amount > 0 && txt.includes('ATM Withdraw')) {
        terminalAggregation[tid].atm_withdraw_nonzero_count++;
      }
    }
  }

  setProgress('atm', 70, 'Waiting for opening balances...');

  // Show opening balance modal
  const openingBalances = await showOpeningBalanceModal();
  if (openingBalances === null) {
    showStatus('atmStatus', 'Opening balance entry was canceled.', 'warning');
    hideProgress('atm');
    return;
  }

  setProgress('atm', 85, 'Generating report...');

  // Calculate summaries
  const summaryData = [];
  TRANSACTION_CATEGORIES.forEach(cat => {
    if (sheets[cat].length > 0 || ['ATM_Withdraw_Fee_NonZero', 'ATM_Transfer_CardToAccount'].includes(cat)) {
      summaryData.push(calculateSheetSummary(sheets[cat], amountIdx, feeIdx, cardIdx, cat));
    }
  });

  // Store all data for export and display
  atmProcessedData = {
    header, sheets, summaryData, terminalAggregation, terminalTransactions,
    openingBalances, posEcomCount, posEcomAmount, issuerCommissionCount,
    issuerCommission, cardToCardCount, cardToCardAmount
  };

  setProgress('atm', 100, 'Complete!');
  renderAtmResults();
  showStatus('atmStatus', 'Processing complete. Review results below.', 'success');
  showToast('ATM transaction file processed successfully', 'success');
}

function renderAtmResults() {
  if (!atmProcessedData) return;
  const d = atmProcessedData;

  // Summary grid
  const grid = document.getElementById('atmSummaryGrid');
  const totalTrans = d.summaryData.reduce((s, x) => s + x.count, 0);
  const totalAmt = d.summaryData.reduce((s, x) => s + x.totalAmount, 0);
  const totalFee = d.summaryData.reduce((s, x) => s + x.totalFee, 0);
  grid.innerHTML = `
    <div class="stat-card blue"><div class="stat-label">Total Transactions</div><div class="stat-value">${totalTrans.toLocaleString()}</div><div class="stat-sub">Across all categories</div></div>
    <div class="stat-card gold"><div class="stat-label">Total Amount</div><div class="stat-value">${formatNumber(totalAmt)}</div><div class="stat-sub">MMK</div></div>
    <div class="stat-card green"><div class="stat-label">Total Fee</div><div class="stat-value">${formatNumber(totalFee)}</div><div class="stat-sub">MMK</div></div>
    <div class="stat-card red"><div class="stat-label">Net Transaction</div><div class="stat-value">${formatNumber(totalAmt - totalFee)}</div><div class="stat-sub">Amount - Fee</div></div>
  `;
  grid.style.display = 'grid';

  // Terminal details table
  const tbody = document.getElementById('atmTerminalBody');
  let html = '';
  let sn = 1;
  TERMINAL_IDS.forEach(tid => {
    const agg = d.terminalAggregation[tid];
    const opening = d.openingBalances[tid] || 0;
    const trans = agg.atm_withdraw_nonzero_count;
    const issuer = agg.atm_issuer_withdraw_amount + agg.atm_card_950316_amount;
    const acquirer = agg.atm_withdraw_fee_nonzero_trx_amount - agg.atm_withdraw_fee_nonzero_amount;
    const balance = opening - issuer - acquirer;
    const loc = TERMINAL_LOCATIONS[tid] || '';
    html += `<tr>
      <td>${sn++}</td>
      <td class="text-primary">${tid}</td>
      <td>${loc}</td>
      <td class="num">${formatNumber(opening)}</td>
      <td class="num">${trans}</td>
      <td class="num">${formatNumber(issuer)}</td>
      <td class="num">${formatNumber(acquirer)}</td>
      <td class="num ${balance < 0 ? 'text-red' : 'text-green'}">${formatNumber(balance)}</td>
      <td></td>
    </tr>`;
  });

  // POS & Ecom row
  html += `<tr class="highlight-row">
    <td></td><td colspan="2">POS & Ecommerce Amount</td><td></td>
    <td class="num">${d.posEcomCount}</td><td></td><td></td>
    <td class="num text-accent">${formatNumber(d.posEcomAmount)}</td><td></td>
  </tr>`;
  // Issuer Off us row
  html += `<tr class="highlight-row">
    <td></td><td colspan="2">Issuer Off us & Commission (Paid)</td><td></td>
    <td class="num">${d.issuerCommissionCount}</td><td></td><td></td>
    <td class="num text-accent">${formatNumber(d.issuerCommission)}</td><td></td>
  </tr>`;

  tbody.innerHTML = html;

  // Summary table
  const sumBody = document.getElementById('atmSummaryBody');
  let sumHtml = '';
  d.summaryData.forEach((s, i) => {
    const isHighlight = ['E2','E4','E5','E6','E7'].some(c => {
      const row = parseInt(c.slice(1)) - 1;
      const col = c[0];
      return (col === 'E' && i === row) || (col === 'G' && i === 2 && ['H3'].includes(c));
    });
    sumHtml += `<tr>
      <td><span class="badge badge-blue">${s.sheetName}</span></td>
      <td class="num">${s.count}</td>
      <td class="num">${formatNumber(s.totalAmount)}</td>
      <td class="num">${formatNumber(s.totalFee)}</td>
      <td class="num text-green" style="font-weight:600">${formatNumber(s.transactionAmount)}</td>
      <td class="num">${s.totalUseCard}</td>
      <td class="num">${formatNumber(s.total1Percent)}</td>
      <td class="num">${formatNumber(s.total1Div2100)}</td>
      <td class="num">${formatNumber(s.totalDiv2100)}</td>
    </tr>`;
  });
  sumBody.innerHTML = sumHtml;

  // Sub-tabs for transaction details
  const subTabs = document.getElementById('atmSubTabs');
  let tabHtml = '';
  TRANSACTION_CATEGORIES.forEach((cat, i) => {
    const count = (d.sheets[cat] || []).length;
    tabHtml += `<button class="sub-tab ${i === 0 ? 'active' : ''}" onclick="showAtmDetail('${cat}', this)">${cat.replace(/_/g,' ')} (${count})</button>`;
  });
  subTabs.innerHTML = tabHtml;

  // Show first category
  if (TRANSACTION_CATEGORIES.length > 0) {
    showAtmDetail(TRANSACTION_CATEGORIES[0]);
  }

  document.getElementById('atmResults').style.display = 'block';
}

function showAtmDetail(category, tabBtn) {
  if (!atmProcessedData) return;
  if (tabBtn) {
    document.querySelectorAll('#atmSubTabs .sub-tab').forEach(t => t.classList.remove('active'));
    tabBtn.classList.add('active');
  }

  const container = document.getElementById('atmDetailTableContainer');
  const rows = atmProcessedData.sheets[category] || [];
  const header = atmProcessedData.header;

  if (rows.length === 0) {
    container.innerHTML = `<div class="empty-state"><i class="fas fa-inbox"></i><h3>No transactions</h3><p>No transactions found in this category.</p></div>`;
    return;
  }

  let html = '<table><thead><tr>';
  header.forEach(h => { html += `<th>${h || ''}</th>`; });
  html += '</tr></thead><tbody>';
  rows.forEach(r => {
    html += '<tr>';
    r.data.forEach(cell => {
      const val = cell !== null && cell !== undefined ? cell : '';
      html += `<td>${val}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  container.innerHTML = html;
}

// ===================== OPENING BALANCE MODAL =====================
function showOpeningBalanceModal() {
  return new Promise((resolve) => {
    const modal = document.getElementById('openingBalanceModal');
    const textarea = document.getElementById('openingBalanceText');
    textarea.value = '';
    hideStatus('openingBalanceStatus');
    modal.classList.add('visible');

    // Store resolve function for later use
    modal._resolve = resolve;

    // Focus textarea
    setTimeout(() => textarea.focus(), 300);
  });
}

function closeOpeningBalanceModal() {
  const modal = document.getElementById('openingBalanceModal');
  modal.classList.remove('visible');
  if (modal._resolve) {
    modal._resolve(null);
    modal._resolve = null;
  }
}

function submitOpeningBalances() {
  const modal = document.getElementById('openingBalanceModal');
  const text = document.getElementById('openingBalanceText').value;
  const lines = text.trim().split('\n');

  const balances = {};
  const twoColPattern = /^([^\t]+)\t([^\t]+)$/;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    const match = trimmed.match(twoColPattern);
    if (match) {
      let tid = match[1].trim().replace(/[^\w]/g, '');
      const balStr = match[2].trim();
      if (TERMINAL_IDS.includes(tid)) {
        balances[tid] = toFloatSafe(balStr);
      }
    }
  }

  if (Object.keys(balances).length === 0) {
    showStatus('openingBalanceStatus', 'No valid opening balances found. Check the format (Tab-separated: TerminalID\\tBalance).', 'error');
    return;
  }

  const missing = TERMINAL_IDS.filter(t => !(t in balances));
  if (missing.length > 0) {
    // Auto-fill 0 for missing terminals
    missing.forEach(t => balances[t] = 0);
    showStatus('openingBalanceStatus', `Parsed ${Object.keys(balances).length - missing.length} balances. ${missing.length} terminals set to 0.00.`, 'warning');
  } else {
    showStatus('openingBalanceStatus', `All ${Object.keys(balances).length} opening balances parsed successfully.`, 'success');
  }

  // Close modal after short delay to show success message
  setTimeout(() => {
    modal.classList.remove('visible');
    if (modal._resolve) {
      modal._resolve(balances);
      modal._resolve = null;
    }
  }, 600);
}

// ===================== DAILY PROCESSING =====================
async function processDailyData() {
  if (!dailyFileBuffer) return;

  setProgress('daily', 10, 'Parsing Excel file...');

  let parsed;
  try {
    parsed = await parseExcelFile(dailyFileBuffer);
  } catch (e) {
    showStatus('dailyStatus', 'Failed to parse Excel file: ' + e.message, 'error');
    hideProgress('daily');
    return;
  }

  const rows = parsed.jsonData;
  if (rows.length < 2) {
    showStatus('dailyStatus', 'File appears to be empty.', 'error');
    hideProgress('daily');
    return;
  }

  setProgress('daily', 30, 'Finding header columns...');

  // Daily report uses row 2 (index 1) as header
  let header = rows.length >= 2 ? rows[1] : rows[0];
  let startRow = rows.length >= 2 ? 2 : 1;

  const amountIdx = findColumnIndex(header, 'amount');
  const feeIdx = findColumnIndex(header, 'fee');
  const cardIdx = findColumnIndex(header, 'card');
  const refnoIdx = findColumnIndex(header, 'ref');

  if (amountIdx < 0 || feeIdx < 0 || refnoIdx < 0) {
    showStatus('dailyStatus', `Could not find required columns. Detected headers: ${header.map((h,i) => `${i+1}.${h}`).join(', ')}`, 'error');
    hideProgress('daily');
    return;
  }

  setProgress('daily', 50, 'Categorizing transactions...');

  const sheets = processRows(rows, header, startRow, amountIdx, feeIdx, cardIdx, refnoIdx);

  setProgress('daily', 80, 'Calculating summaries...');

  const summaryData = [];
  TRANSACTION_CATEGORIES.forEach(cat => {
    if (sheets[cat].length > 0 || ['ATM_Withdraw_Fee_NonZero', 'ATM_Transfer_CardToAccount'].includes(cat)) {
      summaryData.push(calculateSheetSummary(sheets[cat], amountIdx, feeIdx, cardIdx, cat));
    }
  });

  dailyProcessedData = { header, sheets, summaryData };

  setProgress('daily', 100, 'Complete!');
  renderDailyResults();
  showStatus('dailyStatus', 'Daily report processed successfully.', 'success');
  showToast('Daily report processed successfully', 'success');
}

function renderDailyResults() {
  if (!dailyProcessedData) return;
  const d = dailyProcessedData;

  // Summary table
  const sumBody = document.getElementById('dailySummaryBody');
  let sumHtml = '';
  d.summaryData.forEach(s => {
    sumHtml += `<tr>
      <td><span class="badge badge-blue">${s.sheetName}</span></td>
      <td class="num">${s.count}</td>
      <td class="num">${formatNumber(s.totalAmount)}</td>
      <td class="num">${formatNumber(s.totalFee)}</td>
      <td class="num text-green" style="font-weight:600">${formatNumber(s.transactionAmount)}</td>
      <td class="num">${s.totalUseCard}</td>
      <td class="num">${formatNumber(s.total1Percent)}</td>
      <td class="num">${formatNumber(s.total1Div2100)}</td>
      <td class="num">${formatNumber(s.totalDiv2100)}</td>
    </tr>`;
  });
  sumBody.innerHTML = sumHtml;

  // Sub-tabs
  const subTabs = document.getElementById('dailySubTabs');
  let tabHtml = '';
  TRANSACTION_CATEGORIES.forEach((cat, i) => {
    const count = (d.sheets[cat] || []).length;
    tabHtml += `<button class="sub-tab ${i === 0 ? 'active' : ''}" onclick="showDailyDetail('${cat}', this)">${cat.replace(/_/g,' ')} (${count})</button>`;
  });
  subTabs.innerHTML = tabHtml;

  if (TRANSACTION_CATEGORIES.length > 0) showDailyDetail(TRANSACTION_CATEGORIES[0]);

  document.getElementById('dailyResults').style.display = 'block';
}

function showDailyDetail(category, tabBtn) {
  if (!dailyProcessedData) return;
  if (tabBtn) {
    document.querySelectorAll('#dailySubTabs .sub-tab').forEach(t => t.classList.remove('active'));
    tabBtn.classList.add('active');
  }

  const container = document.getElementById('dailyDetailTableContainer');
  const rows = dailyProcessedData.sheets[category] || [];
  const header = dailyProcessedData.header;

  if (rows.length === 0) {
    container.innerHTML = `<div class="empty-state"><i class="fas fa-inbox"></i><h3>No transactions</h3><p>No transactions in this category.</p></div>`;
    return;
  }

  let html = '<table><thead><tr>';
  header.forEach(h => { html += `<th>${h || ''}</th>`; });
  html += '</tr></thead><tbody>';
  rows.forEach(r => {
    html += '<tr>';
    r.data.forEach(cell => {
      html += `<td>${cell !== null && cell !== undefined ? cell : ''}</td>`;
    });
    html += '</tr>';
  });
  html += '</tbody></table>';
  container.innerHTML = html;
}

// ===================== ICBS PROCESSING =====================
async function processIcbsData() {
  if (!icbsFileBuffer) return;

  showStatus('icbsStatus', 'Parsing file...', 'info');

  let parsed;
  try {
    parsed = await parseExcelFile(icbsFileBuffer);
  } catch (e) {
    showStatus('icbsStatus', 'Failed to parse file: ' + e.message, 'error');
    return;
  }

  const rows = parsed.jsonData;
  if (rows.length < 2) {
    showStatus('icbsStatus', 'File is empty or has no data rows.', 'error');
    return;
  }

  // Find header
  let header = rows[0];
  let startRow = 1;
  for (let r = 0; r < Math.min(3, rows.length); r++) {
    const txt = rows[r].map(h => String(h || '').toLowerCase()).join(' ');
    if (txt.includes('account') || txt.includes('debit') || txt.includes('credit') || txt.includes('trans')) {
      header = rows[r];
      startRow = r + 1;
      break;
    }
  }

  const dataRows = rows.slice(startRow).filter(r => r.some(c => c !== ''));

  icbsProcessedData = { header, rows: dataRows, allRows: dataRows };

  renderIcbsResults();
  showStatus('icbsStatus', `Loaded ${dataRows.length} transactions successfully.`, 'success');
  showToast(`ICBS: ${dataRows.length} transactions loaded`, 'success');
}

function renderIcbsResults() {
  if (!icbsProcessedData) return;
  const d = icbsProcessedData;
  const header = d.header;

  // Summary grid
  const grid = document.getElementById('icbsSummaryGrid');
  let totalDebit = 0, totalCredit = 0, debitCount = 0, creditCount = 0;
  d.rows.forEach(r => {
    // Try to find debit/credit columns
    const debitIdx = findColumnIndex(header, 'debit');
    const creditIdx = findColumnIndex(header, 'credit');
    const ttIdx = findColumnIndex(header, 't/t');
    if (debitIdx >= 0) {
      const dv = toFloatSafe(r[debitIdx]);
      if (dv > 0) { totalDebit += dv; debitCount++; }
    }
    if (creditIdx >= 0) {
      const cv = toFloatSafe(r[creditIdx]);
      if (cv > 0) { totalCredit += cv; creditCount++; }
    }
  });

  grid.innerHTML = `
    <div class="stat-card blue"><div class="stat-label">Total Rows</div><div class="stat-value">${d.rows.length.toLocaleString()}</div></div>
    <div class="stat-card red"><div class="stat-label">Total Debit</div><div class="stat-value">${formatNumber(totalDebit)}</div><div class="stat-sub">${debitCount} transactions</div></div>
    <div class="stat-card green"><div class="stat-label">Total Credit</div><div class="stat-value">${formatNumber(totalCredit)}</div><div class="stat-sub">${creditCount} transactions</div></div>
    <div class="stat-card gold"><div class="stat-label">Net</div><div class="stat-value">${formatNumber(totalCredit - totalDebit)}</div><div class="stat-sub">Credit - Debit</div></div>
  `;
  grid.style.display = 'grid';

  // Table
  document.getElementById('icbsTableHead').innerHTML = '<tr>' + header.map(h => `<th>${h || ''}</th>`).join('') + '</tr>';
  renderIcbsTableBody(d.rows);

  document.getElementById('icbsFilterCard').style.display = 'block';
  document.getElementById('icbsTableCard').style.display = 'block';
  document.getElementById('icbsExportBtn').disabled = false;
}

function renderIcbsTableBody(rows) {
  if (!icbsProcessedData) return;
  const tbody = document.getElementById('icbsTableBody');
  const header = icbsProcessedData.header;
  const debitIdx = findColumnIndex(header, 'debit');
  const creditIdx = findColumnIndex(header, 'credit');

  let html = '';
  rows.forEach(r => {
    html += '<tr>';
    r.forEach((cell, i) => {
      let cls = '';
      if (i === debitIdx && toFloatSafe(cell) > 0) cls = 'text-red num';
      else if (i === creditIdx && toFloatSafe(cell) > 0) cls = 'text-green num';
      else if (i === debitIdx || i === creditIdx) cls = 'num';
      html += `<td class="${cls}">${cell !== null && cell !== undefined ? cell : ''}</td>`;
    });
    html += '</tr>';
  });
  tbody.innerHTML = html;
  document.getElementById('icbsRowCount').textContent = rows.length + ' rows';
}

function applyIcbsFilters() {
  if (!icbsProcessedData) return;
  const d = icbsProcessedData;
  const accountFilter = document.getElementById('icbsFilterAccount').value.trim().toLowerCase();
  const descFilter = document.getElementById('icbsFilterDesc').value.trim().toLowerCase();
  const typeFilter = document.getElementById('icbsFilterType').value;
  const terminalFilter = document.getElementById('icbsFilterTerminal').value.trim().toLowerCase();

  const header = d.header;
  const acctIdx = findColumnIndex(header, 'account');
  const descIdx = findColumnIndex(header, 'description');
  const ttIdx = findColumnIndex(header, 't/t');
  const termIdx = findColumnIndex(header, 'terminal');

  const filtered = d.allRows.filter(r => {
    if (accountFilter && acctIdx >= 0) {
      if (!String(r[acctIdx] || '').toLowerCase().includes(accountFilter)) return false;
    }
    if (descFilter && descIdx >= 0) {
      if (!String(r[descIdx] || '').toLowerCase().includes(descFilter)) return false;
    }
    if (typeFilter && ttIdx >= 0) {
      if (String(r[ttIdx] || '').trim().toUpperCase() !== typeFilter) return false;
    }
    if (terminalFilter && termIdx >= 0) {
      if (!String(r[termIdx] || '').toLowerCase().includes(terminalFilter)) return false;
    }
    return true;
  });

  d.rows = filtered;
  renderIcbsTableBody(filtered);
}

function clearIcbsFilters() {
  document.getElementById('icbsFilterAccount').value = '';
  document.getElementById('icbsFilterDesc').value = '';
  document.getElementById('icbsFilterType').value = '';
  document.getElementById('icbsFilterTerminal').value = '';
  if (icbsProcessedData) {
    icbsProcessedData.rows = icbsProcessedData.allRows;
    renderIcbsTableBody(icbsProcessedData.allRows);
  }
}

// ===================== CONVERTER =====================
async function convertFile() {
  if (!converterFileBuffer) return;

  showStatus('converterStatus', 'Converting .xls to .xlsx...', 'info');

  try {
    const data = new Uint8Array(converterFileBuffer);
    const wb = XLSX.read(data, { type: 'array' });
    const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = document.getElementById('converterFileName').textContent.replace(/\.xls$/i, '_converted.xlsx');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showStatus('converterStatus', 'File converted and downloaded successfully.', 'success');
    showToast('File converted successfully', 'success');
  } catch (e) {
    showStatus('converterStatus', 'Conversion failed: ' + e.message, 'error');
  }
}

// ===================== EXPORT FUNCTIONS =====================
function exportAtmReport() {
  if (!atmProcessedData) return;
  const d = atmProcessedData;

  const wb = XLSX.utils.book_new();

  // Terminal Details sheet
  const terminalData = [];
  terminalData.push(['S/N', 'Terminal ID', 'Location', 'Opening Balance', 'Trans', 'Issuer', 'Acquirer', 'Balance', 'Remark']);
  let sn = 1;
  TERMINAL_IDS.forEach(tid => {
    const agg = d.terminalAggregation[tid];
    const opening = d.openingBalances[tid] || 0;
    const trans = agg.atm_withdraw_nonzero_count;
    const issuer = agg.atm_issuer_withdraw_amount + agg.atm_card_950316_amount;
    const acquirer = agg.atm_withdraw_fee_nonzero_trx_amount - agg.atm_withdraw_fee_nonzero_amount;
    const balance = opening - issuer - acquirer;
    terminalData.push([sn++, tid, TERMINAL_LOCATIONS[tid] || '', opening, trans, issuer, acquirer, balance, '']);
  });
  terminalData.push(['', 'POS & Ecommerce Amount', '', '', d.posEcomCount, '', '', d.posEcomAmount, '']);
  terminalData.push(['', 'Issuer Off us & Commission (Paid)', '', '', d.issuerCommissionCount, '', '', d.issuerCommission, '']);
  const wsTerminal = XLSX.utils.aoa_to_sheet(terminalData);
  XLSX.utils.book_append_sheet(wb, wsTerminal, 'Terminal Details');

  // Category sheets
  TRANSACTION_CATEGORIES.forEach(cat => {
    const rows = d.sheets[cat] || [];
    const sheetData = [d.header];
    rows.forEach(r => sheetData.push(r.data));

    // Add summary rows
    const summary = d.summaryData.find(s => s.sheetName === cat);
    if (summary) {
      sheetData.push([]);
      sheetData.push(['Total Count', summary.count]);
      sheetData.push(['Total Amount', summary.totalAmount]);
      sheetData.push(['Total Fee', summary.totalFee]);
      sheetData.push(['Transaction Amount', summary.transactionAmount]);
      sheetData.push(['Total Use Card', summary.totalUseCard]);
      sheetData.push(['Total 1%', summary.total1Percent]);
      sheetData.push(['Total 1% / 2100', summary.total1Div2100]);
      sheetData.push(['Total / 2100', summary.totalDiv2100]);
    }

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(wb, ws, cat);
  });

  // Summary sheet
  const summaryData = [['Sheet Name', 'Total Count', 'Total Amount', 'Total Fee', 'Total Transaction', 'Total Use Card', 'Total 1%', 'Total 1% / 2100', 'Total / 2100']];
  d.summaryData.forEach(s => {
    summaryData.push([s.sheetName, s.count, s.totalAmount, s.totalFee, s.transactionAmount, s.totalUseCard, s.total1Percent, s.total1Div2100, s.totalDiv2100]);
  });
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  // Terminal transaction detail sheets
  TERMINAL_IDS.forEach(tid => {
    const tRows = d.terminalTransactions[tid];
    if (tRows.length === 0) return;
    const sheetData = [d.header];
    tRows.forEach(r => sheetData.push(r));
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(wb, ws, `T_${tid}`);
  });

  // Download
  const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `ATM_Report_${new Date().toISOString().slice(0,10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('ATM report exported successfully', 'success');
}

function exportDailyReport() {
  if (!dailyProcessedData) return;
  const d = dailyProcessedData;

  const wb = XLSX.utils.book_new();

  TRANSACTION_CATEGORIES.forEach(cat => {
    const rows = d.sheets[cat] || [];
    const sheetData = [d.header];
    rows.forEach(r => sheetData.push(r.data));

    const summary = d.summaryData.find(s => s.sheetName === cat);
    if (summary) {
      sheetData.push([]);
      sheetData.push(['Total Count', summary.count]);
      sheetData.push(['Total Amount', summary.totalAmount]);
      sheetData.push(['Total Fee', summary.totalFee]);
      sheetData.push(['Transaction Amount', summary.transactionAmount]);
      sheetData.push(['Total Use Card', summary.totalUseCard]);
      sheetData.push(['Total 1%', summary.total1Percent]);
      sheetData.append(['Total 1% / 2100', summary.total1Div2100]);
      sheetData.push(['Total / 2100', summary.totalDiv2100]);
    }

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(wb, ws, cat);
  });

  // Summary sheet
  const summaryData = [['Sheet Name', 'Total Count', 'Total Amount', 'Total Fee', 'Total Transaction', 'Total Use Card', 'Total 1%', 'Total 1% / 2100', 'Total / 2100']];
  d.summaryData.forEach(s => {
    summaryData.push([s.sheetName, s.count, s.totalAmount, s.totalFee, s.transactionAmount, s.totalUseCard, s.total1Percent, s.total1Div2100, s.totalDiv2100]);
  });
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Daily_Report_${new Date().toISOString().slice(0,10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('Daily report exported successfully', 'success');
}

function exportIcbsReport() {
  if (!icbsProcessedData) return;
  const d = icbsProcessedData;

  const wb = XLSX.utils.book_new();
  const sheetData = [d.header];
  d.rows.forEach(r => sheetData.push(r));
  const ws = XLSX.utils.aoa_to_sheet(sheetData);
  XLSX.utils.book_append_sheet(wb, ws, 'ICBS Data');

  const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `ICBS_Report_${new Date().toISOString().slice(0,10)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast('ICBS report exported successfully', 'success');
}

// ===================== KEYBOARD SHORTCUTS =====================
document.addEventListener('keydown', (e) => {
  // Escape to close modal
  if (e.key === 'Escape') {
    const modal = document.getElementById('openingBalanceModal');
    if (modal.classList.contains('visible')) {
      closeOpeningBalanceModal();
    }
  }
});

// Close modal on overlay click
document.getElementById('openingBalanceModal').addEventListener('click', (e) => {
  if (e.target === e.currentTarget) closeOpeningBalanceModal();
});
</script>
</body>
</html>
