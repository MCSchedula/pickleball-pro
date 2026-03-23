<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Planificateur Pickleball Pro - Web App</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="{{ url_for('static', filename='style.css') }}">
</head>
<body>
    <div class="app-container">
        <!-- Sidebar -->
        <aside class="sidebar">
            <div class="logo">
                <h1>🏓 Pickleball Pro</h1>
                <p>Application Web Python</p>
            </div>

            <nav class="nav-menu">
                <div class="nav-section">
                    <div class="nav-section-title">Configuration</div>
                    <div class="nav-item active" data-view="dashboard">
                        <span class="nav-item-icon">📊</span>
                        <span>Tableau de bord</span>
                    </div>
                    <div class="nav-item" data-view="players">
                        <span class="nav-item-icon">👥</span>
                        <span>Joueurs</span>
                    </div>
                    <div class="nav-item" data-view="events">
                        <span class="nav-item-icon">📅</span>
                        <span>Événements</span>
                    </div>
                </div>

                <div class="nav-section">
                    <div class="nav-section-title">Cédules</div>
                    <div class="nav-item" data-view="generate">
                        <span class="nav-item-icon">⚡</span>
                        <span>Générer Cédule</span>
                    </div>
                    <div class="nav-item" data-view="results">
                        <span class="nav-item-icon">📋</span>
                        <span>Résultats</span>
                    </div>
                </div>

                <div class="nav-section">
                    <div class="nav-section-title">Système</div>
                    <div class="nav-item" data-view="settings">
                        <span class="nav-item-icon">⚙️</span>
                        <span>Paramètres</span>
                    </div>
                </div>
            </nav>
        </aside>

        <!-- Main Content -->
        <main class="main-content">
            <div class="top-bar">
                <div class="top-bar-left">
                    <h2 id="pageTitle">Tableau de bord</h2>
                    <p id="pageSubtitle">Application Web Python/Flask</p>
                </div>
                <div class="top-bar-actions">
                    <button class="btn btn-secondary btn-sm" onclick="loadExcelData()">
                        📥 Charger Excel
                    </button>
                    <button class="btn btn-primary btn-sm" onclick="downloadSchedule()">
                        💾 Télécharger Excel
                    </button>
                </div>
            </div>

            <div class="content-area">
                <!-- Dashboard View -->
                <div id="view-dashboard" class="view-content">
                    <div class="stats-grid">
                        <div class="stat-card">
                            <div class="stat-value" id="stat-players">0</div>
                            <div class="stat-label">Joueurs</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value" id="stat-events">0</div>
                            <div class="stat-label">Événements</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-value" id="stat-schedules">0</div>
                            <div class="stat-label">Cédules générées</div>
                        </div>
                    </div>

                    <div class="card">
                        <div class="card-header">
                            <div>
                                <h3 class="card-title">🌐 Application Web Python/Flask</h3>
                                <p class="card-subtitle">Base de données centralisée</p>
                            </div>
                        </div>
                        <div style="color: var(--text-secondary); line-height: 1.8;">
                            <h4 style="color: var(--primary-light); margin-bottom: 12px;">✨ Avantages de l'application web</h4>
                            <ul style="margin-left: 20px; line-height: 2;">
                                <li>✅ <strong>URL unique</strong> : Tous les responsables accèdent au même site</li>
                                <li>✅ <strong>Base de données partagée</strong> : 374 joueurs synchronisés</li>
                                <li>✅ <strong>Mises à jour automatiques</strong> : Nouveau joueur visible par tous</li>
                                <li>✅ <strong>Algorithme Python</strong> : Plus rapide et performant</li>
                                <li>✅ <strong>Historique complet</strong> : Toutes les cédules sauvegardées</li>
                                <li>✅ <strong>Aucune installation</strong> : Fonctionne dans le navigateur</li>
                            </ul>

                            <h4 style="color: var(--primary-light); margin: 24px 0 12px;">🚀 Guide de démarrage</h4>
                            <p style="margin-bottom: 16px;">
                                <strong>Étape 1 :</strong> Cliquez sur <strong>"📥 Charger Excel"</strong> et sélectionnez :<br>
                                • <code>Base de données - Pickleball.xlsm</code> (tous les joueurs)<br>
                                • <code>Cédules informations.xlsx</code> (événements)
                            </p>

                            <p style="margin-bottom: 16px;">
                                <strong>Étape 2 :</strong> Dans <strong>"⚡ Générer Cédule"</strong> :<br>
                                • Sélectionnez l'événement<br>
                                • Cochez les joueurs participants et drill<br>
                                • Générez la cédule
                            </p>

                            <p>
                                <strong>Étape 3 :</strong> Téléchargez le fichier Excel avec toutes les feuilles formatées
                            </p>
                        </div>
                    </div>
                </div>

                <!-- Generate View -->
                <div id="view-generate" class="view-content hidden">
                    <div class="card">
                        <div class="card-header">
                            <div>
                                <h3 class="card-title">Générer une nouvelle cédule</h3>
                                <p class="card-subtitle">Sélectionnez l'événement et les joueurs</p>
                            </div>
                        </div>

                        <div class="form-group mb-24">
                            <label class="form-label">Événement</label>
                            <select class="form-select" id="generate-event-select">
                                <option value="">-- Sélectionner un événement --</option>
                            </select>
                        </div>

                        <div id="event-details" class="hidden mb-24"></div>

                        <div id="player-selection" class="hidden">
                            <div class="card">
                                <div class="card-header">
                                    <div>
                                        <h3 class="card-title">Sélection des joueurs</h3>
                                        <p class="card-subtitle" id="selected-count">0 joueur sélectionné</p>
                                    </div>
                                    <div style="display: flex; gap: 8px;">
                                        <button class="btn btn-secondary btn-sm" onclick="selectAllPlayers(true, false)">
                                            ✓ Tout sélectionner
                                        </button>
                                        <button class="btn btn-secondary btn-sm" onclick="selectAllPlayers(false, false)">
                                            ✗ Tout désélectionner
                                        </button>
                                    </div>
                                </div>
                                <div class="player-grid" id="player-selection-grid"></div>
                            </div>

                            <div class="mt-16" style="text-align: center;">
                                <button class="btn btn-primary" style="font-size: 16px; padding: 14px 32px;" onclick="generateSchedule()">
                                    ⚡ Générer la cédule
                                </button>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Results View -->
                <div id="view-results" class="view-content hidden">
                    <div id="results-content">
                        <div class="card">
                            <p class="text-muted">Aucune cédule générée.</p>
                        </div>
                    </div>
                </div>

                <!-- Players View -->
                <div id="view-players" class="view-content hidden">
                    <div class="card">
                        <div class="card-header">
                            <h3 class="card-title">Base de données - Joueurs</h3>
                        </div>
                        <div id="players-list"></div>
                    </div>
                </div>

                <!-- Events View -->
                <div id="view-events" class="view-content hidden">
                    <div class="card">
                        <div class="card-header">
                            <h3 class="card-title">Événements configurés</h3>
                        </div>
                        <div id="events-list"></div>
                    </div>
                </div>

                <!-- Settings View -->
                <div id="view-settings" class="view-content hidden">
                    <div class="card">
                        <div class="card-header">
                            <h3 class="card-title">Paramètres du système</h3>
                            <p class="card-subtitle">Contraintes de formation d'équipes</p>
                        </div>
                        
                        <div class="form-grid">
                            <div class="form-group">
                                <label class="form-label">Maximum fois comme coéquipiers</label>
                                <input type="number" class="form-input" id="setting-max-teammates" 
                                       min="0" max="10" step="1" value="1">
                            </div>
                            
                            <div class="form-group">
                                <label class="form-label">Maximum fois comme adversaires</label>
                                <input type="number" class="form-input" id="setting-max-opponents" 
                                       min="0" max="10" step="1" value="2">
                            </div>
                        </div>
                        
                        <div class="mt-16">
                            <button class="btn btn-primary" onclick="saveSettings()">
                                💾 Sauvegarder les paramètres
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </main>
    </div>

    <script src="{{ url_for('static', filename='app.js') }}"></script>
</body>
</html>
