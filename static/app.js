// Application State
const appState = {
    players: [],
    events: [],
    currentSchedule: null,
    selectedPlayers: [],
    drillPlayers: [],
    settings: {}
};

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    setupNavigation();
    loadData();
});

// Navigation
function setupNavigation() {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', function() {
            const viewName = this.dataset.view;
            if (!viewName) return;

            document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
            this.classList.add('active');

            document.querySelectorAll('.view-content').forEach(el => el.classList.add('hidden'));
            document.getElementById(`view-${viewName}`).classList.remove('hidden');

            const titles = {
                dashboard: ['Tableau de bord', 'Application Web Python/Flask'],
                players: ['Joueurs', 'Base de données'],
                events: ['Événements', 'Configuration'],
                generate: ['Générer Cédule', 'Nouvelle cédule'],
                results: ['Résultats', 'Cédule générée'],
                settings: ['Paramètres', 'Configuration']
            };
            
            if (titles[viewName]) {
                document.getElementById('pageTitle').textContent = titles[viewName][0];
                document.getElementById('pageSubtitle').textContent = titles[viewName][1];
            }

            if (viewName === 'generate') initGenerateView();
            if (viewName === 'players') loadPlayersView();
            if (viewName === 'events') loadEventsView();
            if (viewName === 'settings') loadSettingsView();
        });
    });
}

// Load Data from API
async function loadData() {
    try {
        const [playersRes, eventsRes, settingsRes] = await Promise.all([
            fetch('/api/players'),
            fetch('/api/events'),
            fetch('/api/settings')
        ]);

        appState.players = await playersRes.json();
        appState.events = await eventsRes.json();
        appState.settings = await settingsRes.json();

        updateDashboard();
    } catch (error) {
        console.error('Error loading data:', error);
        showToast('Erreur de chargement', 'error');
    }
}

function updateDashboard() {
    document.getElementById('stat-players').textContent = appState.players.length;
    document.getElementById('stat-events').textContent = appState.events.length;
}

// Excel Upload
function loadExcelData() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xlsm';
    input.onchange = handleFileUpload;
    input.click();
}

async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    try {
        const res = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        const result = await res.json();
        showToast(`Chargé: ${result.players} joueurs, ${result.events} événements, ${result.selected} sélectionnés, ${result.drill} drill`);
        
        await loadData();
    } catch (error) {
        showToast('Erreur: ' + error.message, 'error');
    }
}

// Generate View
function initGenerateView() {
    const select = document.getElementById('generate-event-select');
    select.innerHTML = '<option value="">-- Sélectionner un événement --</option>';
    
    appState.events.forEach(e => {
        select.innerHTML += `<option value="${e.id}">${e.name} - ${e.day}</option>`;
    });

    select.onchange = function() {
        if (this.value) {
            showEventDetails(parseInt(this.value));
            showPlayerSelection();
        } else {
            document.getElementById('event-details').classList.add('hidden');
            document.getElementById('player-selection').classList.add('hidden');
        }
    };
}

function showEventDetails(eventId) {
    const event = appState.events.find(e => e.id === eventId);
    if (!event) return;

    document.getElementById('event-details').innerHTML = `
        <div class="card" style="background: var(--bg-elevated);">
            <p style="color: var(--text-secondary); margin-bottom: 12px;">
                <strong>${event.name}</strong> - ${event.day}
            </p>
            <p style="color: var(--text-muted); font-size: 14px;">
                🕐 ${event.startTime} - ${event.endTime} &nbsp;|&nbsp;
                ⚙️ Drill: ${event.drillMinutes} min &nbsp;|&nbsp;
                ⏱ Périodes: ${event.periodDuration} min
            </p>
        </div>
    `;
    document.getElementById('event-details').classList.remove('hidden');
}

function showPlayerSelection() {
    appState.selectedPlayers = appState.players
        .filter(p => p.selected)
        .map(p => p.id);

    appState.drillPlayers = appState.players
        .filter(p => p.drill)
        .map(p => p.id);

    const activePlayers = appState.players.filter(p => p.status === 'Actif' || !p.status);
    const grid = document.getElementById('player-selection-grid');
    
    grid.innerHTML = activePlayers.map(p => {
        const isSelected = appState.selectedPlayers.includes(p.id);
        const isDrill = appState.drillPlayers.includes(p.id);
        
        return `
        <div class="player-card ${isSelected ? 'selected' : ''}" data-id="${p.id}">
            <input type="checkbox" class="player-checkbox" ${isSelected ? 'checked' : ''} 
                   onchange="handlePlayerCheckbox(${p.id}, this.checked)">
            <div class="player-info" onclick="toggleCardFromInfo(${p.id})">
                <div class="player-name">${p.fullName}</div>
                <div class="player-meta">
                    ${p.gender} · <span class="level-badge">${p.level}</span>
                    <label style="margin-left: 12px;" onclick="event.stopPropagation();">
                        <input type="checkbox" class="drill-checkbox" ${isDrill ? 'checked' : ''} 
                               onchange="toggleDrill(${p.id})">
                        Drill
                    </label>
                </div>
            </div>
        </div>`;
    }).join('');

    document.getElementById('player-selection').classList.remove('hidden');
    updateSelectedCount();
}

function handlePlayerCheckbox(id, checked) {
    const card = document.querySelector(`.player-card[data-id="${id}"]`);
    
    if (checked) {
        card.classList.add('selected');
        if (!appState.selectedPlayers.includes(id)) {
            appState.selectedPlayers.push(id);
        }
    } else {
        card.classList.remove('selected');
        appState.selectedPlayers = appState.selectedPlayers.filter(pid => pid !== id);
        appState.drillPlayers = appState.drillPlayers.filter(pid => pid !== id);
        const drillCb = card.querySelector('.drill-checkbox');
        if (drillCb) drillCb.checked = false;
    }
    
    updateSelectedCount();
}

function toggleCardFromInfo(id) {
    const card = document.querySelector(`.player-card[data-id="${id}"]`);
    const checkbox = card.querySelector('.player-checkbox');
    checkbox.checked = !checkbox.checked;
    handlePlayerCheckbox(id, checkbox.checked);
}

function toggleDrill(id) {
    if (appState.drillPlayers.includes(id)) {
        appState.drillPlayers = appState.drillPlayers.filter(pid => pid !== id);
    } else {
        appState.drillPlayers.push(id);
    }
    updateSelectedCount();
}

function selectAllPlayers(select, drillOnly) {
    document.querySelectorAll('.player-card').forEach(card => {
        const id = parseInt(card.dataset.id);
        
        if (drillOnly) {
            if (appState.selectedPlayers.includes(id)) {
                const drillCb = card.querySelector('.drill-checkbox');
                drillCb.checked = select;
                if (select && !appState.drillPlayers.includes(id)) {
                    appState.drillPlayers.push(id);
                } else if (!select) {
                    appState.drillPlayers = appState.drillPlayers.filter(pid => pid !== id);
                }
            }
        } else {
            const checkbox = card.querySelector('.player-checkbox');
            if (select) {
                card.classList.add('selected');
                checkbox.checked = true;
                if (!appState.selectedPlayers.includes(id)) {
                    appState.selectedPlayers.push(id);
                }
            } else {
                card.classList.remove('selected');
                checkbox.checked = false;
                appState.selectedPlayers = appState.selectedPlayers.filter(pid => pid !== id);
                appState.drillPlayers = appState.drillPlayers.filter(pid => pid !== id);
                card.querySelector('.drill-checkbox').checked = false;
            }
        }
    });
    updateSelectedCount();
}

function updateSelectedCount() {
    document.getElementById('selected-count').textContent = 
        `${appState.selectedPlayers.length} joueur${appState.selectedPlayers.length > 1 ? 's' : ''} sélectionné${appState.selectedPlayers.length > 1 ? 's' : ''}`;
}

// Generate Schedule
async function downloadSchedule() {
    if (!appState.currentSchedule) {
        const saved = localStorage.getItem('currentSchedule');
        if (saved) {
            appState.currentSchedule = JSON.parse(saved);
        }
    }

    if (!appState.currentSchedule) {
        showToast('Aucune cédule à télécharger', 'error');
        return;
    }

    try {
        const res = await fetch('/api/export-excel', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(appState.currentSchedule)
        });

        if (!res.ok) {
            throw new Error('Erreur lors de la génération du fichier Excel');
        }

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = 'Cedule_de_la_journee.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();

        window.URL.revokeObjectURL(url);

        showToast('Téléchargement Excel démarré');
    } catch (error) {
        console.error(error);
        showToast('Erreur: ' + error.message, 'error');
    }
}

function displayResults(schedule) {
    let html = '<div class="card">';
    html += '<div class="card-header">';
    html += `<div><h3 class="card-title">Cédule - ${schedule.event.name}</h3></div>`;
    html += '</div>';

    schedule.periods.forEach(period => {
        html += '<div style="margin-bottom: 24px;">';
        html += `<div style="padding: 12px; background: var(--bg-elevated); border-radius: 8px; margin-bottom: 12px;">`;
        html += `<strong style="color: var(--accent);">${period.time}</strong> - ${period.name}`;
        if (period.sitting && period.sitting.length > 0) {
            html += ` <span style="color: var(--text-muted);">⏸️ Pause: ${period.sitting.map(p => p.fullName).join(', ')}</span>`;
        }
        html += '</div>';

        period.courts.forEach(court => {
            html += `<div style="padding: 12px; background: var(--bg-card); border-left: 4px solid var(--primary); margin-bottom: 8px; border-radius: 8px;">`;
            html += `<strong>Terrain ${court.number}</strong><br>`;
            html += `Côté A: ${court.sideA.player1.fullName}, ${court.sideA.player2.fullName}<br>`;
            html += `Côté B: ${court.sideB.player1.fullName}, ${court.sideB.player2.fullName}`;
            html += '</div>';
        });

        html += '</div>';
    });

    html += '</div>';
    document.getElementById('results-content').innerHTML = html;
}

// Settings
async function loadSettingsView() {
    document.getElementById('setting-max-teammates').value = appState.settings.maxTeammates || 1;
    document.getElementById('setting-max-opponents').value = appState.settings.maxOpponents || 2;
}

async function saveSettings() {
    const settings = {
        maxTeammates: parseInt(document.getElementById('setting-max-teammates').value),
        maxOpponents: parseInt(document.getElementById('setting-max-opponents').value)
    };

    try {
        await fetch('/api/settings', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(settings)
        });

        appState.settings = settings;
        showToast('Paramètres sauvegardés!');
    } catch (error) {
        showToast('Erreur: ' + error.message, 'error');
    }
}

// Players View
function loadPlayersView() {
    const html = `
        <table style="width: 100%; color: var(--text-secondary);">
            <thead>
                <tr style="border-bottom: 1px solid var(--border);">
                    <th style="text-align: left; padding: 12px;">Nom</th>
                    <th style="text-align: left; padding: 12px;">Genre</th>
                    <th style="text-align: left; padding: 12px;">Niveau</th>
                    <th style="text-align: left; padding: 12px;">Statut</th>
                </tr>
            </thead>
            <tbody>
                ${appState.players.map(p => `
                    <tr style="border-bottom: 1px solid var(--border);">
                        <td style="padding: 12px;">${p.fullName}</td>
                        <td style="padding: 12px;">${p.gender}</td>
                        <td style="padding: 12px;"><span class="level-badge">${p.level}</span></td>
                        <td style="padding: 12px;">${p.status}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
    document.getElementById('players-list').innerHTML = html;
}

// Events View
function loadEventsView() {
    const html = appState.events.map(e => `
        <div class="card" style="background: var(--bg-elevated); margin-bottom: 12px;">
            <h4 style="margin-bottom: 8px;">${e.name}</h4>
            <p style="color: var(--text-muted); font-size: 14px;">
                ${e.day} • ${e.startTime} - ${e.endTime} • Drill: ${e.drillMinutes} min
            </p>
        </div>
    `).join('');
    document.getElementById('events-list').innerHTML = html;
}

// Toast
function showToast(message, type = 'success') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}
