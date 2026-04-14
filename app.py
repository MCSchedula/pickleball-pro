from flask import Flask, render_template, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from datetime import datetime, timedelta
import json
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import Counter
import random

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///pickleball.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'pickleball-dsp-secret-key-2026'
db = SQLAlchemy(app)
migrate = Migrate(app, db)

# ==================== MODELS ====================

class Player(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    first_name = db.Column(db.String(100), nullable=False)
    last_name = db.Column(db.String(100), nullable=False)
    full_name = db.Column(db.String(200), nullable=False)
    gender = db.Column(db.String(1), default='M')
    level = db.Column(db.Float, default=3.5)
    email = db.Column(db.String(200))
    status = db.Column(db.String(50), default='Actif')
    selected = db.Column(db.Boolean, default=False)
    drill = db.Column(db.Boolean, default=False)
    
    def to_dict(self):
        return {
            'id': self.id,
            'firstName': self.first_name,
            'lastName': self.last_name,
            'fullName': self.full_name,
            'gender': self.gender,
            'level': self.level,
            'email': self.email,
            'status': self.status,
            'selected': self.selected,
            'drill': self.drill
        }

class Event(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    client = db.Column(db.String(200))
    day = db.Column(db.String(50))
    start_time = db.Column(db.String(10))
    end_time = db.Column(db.String(10))
    drill_minutes = db.Column(db.Integer, default=0)
    period_duration = db.Column(db.Integer, default=20)
    cost = db.Column(db.Float, default=0.0)
    
    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'client': self.client,
            'day': self.day,
            'startTime': self.start_time,
            'endTime': self.end_time,
            'drillMinutes': self.drill_minutes,
            'periodDuration': self.period_duration,
            'cost': self.cost
        }

class Schedule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    event_id = db.Column(db.Integer, db.ForeignKey('event.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    selected_players = db.Column(db.Text)  # JSON array of player IDs
    drill_players = db.Column(db.Text)  # JSON array of player IDs
    schedule_data = db.Column(db.Text)  # JSON of complete schedule
    
    event = db.relationship('Event', backref='schedules')

class Setting(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    key = db.Column(db.String(100), unique=True, nullable=False)
    value = db.Column(db.String(200))
    
    @staticmethod
    def get(key, default=None):
        setting = Setting.query.filter_by(key=key).first()
        return setting.value if setting else default
    
    @staticmethod
    def set(key, value):
        setting = Setting.query.filter_by(key=key).first()
        if setting:
            setting.value = str(value)
        else:
            setting = Setting(key=key, value=str(value))
            db.session.add(setting)
        db.session.commit()

# ==================== ROUTES ====================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/players', methods=['GET'])
def get_players():
    players = Player.query.all()
    return jsonify([p.to_dict() for p in players])

@app.route('/api/events', methods=['GET'])
def get_events():
    events = Event.query.all()
    return jsonify([e.to_dict() for e in events])

@app.route('/api/settings', methods=['GET'])
def get_settings():
    return jsonify({
        'maxTeammates': int(Setting.get('maxTeammates', 1)),
        'maxOpponents': int(Setting.get('maxOpponents', 2)),
        'maxTeamLevelDiff': float(Setting.get('maxTeamLevelDiff', 0.4)),
        'maxMatchLevelDiff': float(Setting.get('maxMatchLevelDiff', 0.49))
    })

@app.route('/api/settings', methods=['POST'])
def update_settings():
    data = request.json
    for key, value in data.items():
        Setting.set(key, value)
    return jsonify({'success': True})

@app.route('/api/reset', methods=['POST'])
def reset_data():
    try:
        # Supprimer toutes les données
        Schedule.query.delete()
        Event.query.delete()
        Player.query.delete()
        Setting.query.delete()
        
        db.session.commit()

        # Recréer les paramètres par défaut
        Setting.set('maxTeammates', 1)
        Setting.set('maxOpponents', 2)
        Setting.set('maxTeamLevelDiff', 0.4)
        Setting.set('maxMatchLevelDiff', 0.49)

        return jsonify({'success': True, 'message': 'Base complètement remise à zéro'})
    
    except Exception as e:
        db.session.rollback()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/upload', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400

    file = request.files['file']
    wb = openpyxl.load_workbook(file, data_only=True)

    result = {'players': 0, 'events': 0, 'selected': 0, 'drill': 0}

    # Import players from "Noms", "Sélection joueurs" or "Membres"
    sheet_name = None
    if 'Noms' in wb.sheetnames:
        sheet_name = 'Noms'
    elif 'Sélection joueurs' in wb.sheetnames:
        sheet_name = 'Sélection joueurs'
    elif 'Membres' in wb.sheetnames:
        sheet_name = 'Membres'

    if sheet_name:
        ws = wb[sheet_name]
        headers = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]
        print("HEADERS:", headers)

        Player.query.delete()

        def is_filled(value):
            return value is not None and str(value).strip() != ''

        for row in ws.iter_rows(min_row=2, values_only=True):
            first_name = ''
            last_name = ''
            full_name = ''

            if 'Prénom' in headers:
                idx = headers.index('Prénom')
                first_name = str(row[idx]).strip() if len(row) > idx and row[idx] else ''

            if 'Nom' in headers:
                idx = headers.index('Nom')
                last_name = str(row[idx]).strip() if len(row) > idx and row[idx] else ''

            if 'Nom complet' in headers:
                idx = headers.index('Nom complet')
                full_name = str(row[idx]).strip() if len(row) > idx and row[idx] else ''
            elif len(row) > 0 and row[0]:
                full_name = str(row[0]).strip()

            if not full_name and (first_name or last_name):
                full_name = f"{first_name} {last_name}".strip()

            if not full_name:
                continue

            gender = 'M'
            if 'Genre' in headers:
                idx = headers.index('Genre')
                gender = str(row[idx]).strip() if len(row) > idx and row[idx] else 'M'

            level = 3.5
            if 'Niveau' in headers:
                idx = headers.index('Niveau')
                try:
                    level = float(row[idx]) if len(row) > idx and row[idx] not in (None, '') else 3.5
                except:
                    level = 3.5

            email = ''
            if 'Courriel' in headers:
                idx = headers.index('Courriel')
                email = str(row[idx]).strip() if len(row) > idx and row[idx] else ''

            status = 'Actif'
            if 'Statut' in headers:
                idx = headers.index('Statut')
                status = str(row[idx]).strip() if len(row) > idx and row[idx] else 'Actif'

            selected = False
            if 'Sélectionner (x)' in headers:
                idx = headers.index('Sélectionner (x)')
                selected = len(row) > idx and is_filled(row[idx])

            drill = False
            if 'Drill (x)' in headers:
                idx = headers.index('Drill (x)')
                drill = len(row) > idx and is_filled(row[idx])

            player = Player(
                first_name=first_name,
                last_name=last_name,
                full_name=full_name,
                gender=gender,
                level=level,
                email=email,
                status=status,
                selected=selected,
                drill=drill
            )

            db.session.add(player)
            result['players'] += 1

            if selected:
                result['selected'] += 1
            if drill:
                result['drill'] += 1

        db.session.commit()

    # Import events
    if 'Événements' in wb.sheetnames:
        ws = wb['Événements']
        headers = [str(cell.value).strip() if cell.value else '' for cell in ws[1]]

        Event.query.delete()

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue

            name = str(row[0]).strip() if row[0] else ''

            client = ''
            if 'Client' in headers:
                idx = headers.index('Client')
                client = str(row[idx]).strip() if len(row) > idx and row[idx] else ''

            day = ''
            if 'Journée' in headers:
                idx = headers.index('Journée')
                day = str(row[idx]).strip() if len(row) > idx and row[idx] else ''

            start_time = '12:00'
            if 'Heure début' in headers:
                idx = headers.index('Heure début')
                if len(row) > idx and row[idx]:
                    val = row[idx]
                    if isinstance(val, datetime):
                        start_time = val.strftime('%H:%M')
                    else:
                        start_time = str(val)

            end_time = '15:00'
            if 'Heure fin' in headers:
                idx = headers.index('Heure fin')
                if len(row) > idx and row[idx]:
                    val = row[idx]
                    if isinstance(val, datetime):
                        end_time = val.strftime('%H:%M')
                    else:
                        end_time = str(val)

            drill_minutes = 0
            if 'Drill en minutes' in headers:
                idx = headers.index('Drill en minutes')
                try:
                    drill_minutes = int(row[idx]) if len(row) > idx and row[idx] else 0
                except:
                    drill_minutes = 0

            period_duration = 20
            if "Durée d'une partie" in headers:
                idx = headers.index("Durée d'une partie")
                try:
                    period_duration = int(row[idx]) if len(row) > idx and row[idx] else 20
                except:
                    period_duration = 20

            cost = 0.0
            if 'Coût pour une cédule' in headers:
                idx = headers.index('Coût pour une cédule')
                try:
                    cost = float(row[idx]) if len(row) > idx and row[idx] else 0.0
                except:
                    cost = 0.0

            event = Event(
                name=name,
                client=client,
                day=day,
                start_time=start_time,
                end_time=end_time,
                drill_minutes=drill_minutes,
                period_duration=period_duration,
                cost=cost
            )

            db.session.add(event)
            result['events'] += 1

        db.session.commit()

    return jsonify(result)

@app.route('/api/generate', methods=['POST'])
def generate_schedule():
    data = request.json
    event_id = data.get('eventId')
    selected_ids = data.get('selectedPlayers', [])
    drill_ids = data.get('drillPlayers', [])
    
    event = Event.query.get(event_id)
    if not event:
        return jsonify({'error': 'Event not found'}), 404
    
    selected_players = [Player.query.get(pid).to_dict() for pid in selected_ids]
    drill_players = [Player.query.get(pid).to_dict() for pid in drill_ids]
    
    settings = {
        'maxTeammates': int(Setting.get('maxTeammates', 1)),
        'maxOpponents': int(Setting.get('maxOpponents', 2))
    }
    
    # Generate schedule using algorithm
    schedule_result = generate_schedule_algorithm(event.to_dict(), selected_players, drill_players, settings)
    
    # Save to database
    schedule = Schedule(
        event_id=event_id,
        selected_players=json.dumps(selected_ids),
        drill_players=json.dumps(drill_ids),
        schedule_data=json.dumps(schedule_result)
    )
    db.session.add(schedule)
    db.session.commit()
    
    return jsonify(schedule_result)

@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    data = request.json
    if not data:
        return jsonify({'error': 'Aucune cédule à exporter'}), 400

    schedule = data

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Cédule de la journée'

    # Styles
    bold_font = Font(bold=True)
    title_font = Font(bold=True, size=14)
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_header = PatternFill(fill_type='solid', fgColor='D9EAF7')

    event = schedule.get('event', {})
    periods = schedule.get('periods', [])

    # En-tête
    ws['A1'] = 'Cédule de la journée'
    ws['A1'].font = title_font

    ws['A2'] = 'Événement'
    ws['B2'] = event.get('name', '')
    ws['A3'] = 'Journée'
    ws['B3'] = event.get('day', '')
    ws['A4'] = 'Heure'
    ws['B4'] = f"{event.get('startTime', '')} à {event.get('endTime', '')}"

    for cell in ['A2', 'A3', 'A4']:
        ws[cell].font = bold_font

    row = 6

    # Colonnes
    ws.cell(row=row, column=1, value='Période')
    ws.cell(row=row, column=2, value='Heure')
    ws.cell(row=row, column=3, value='Terrain')
    ws.cell(row=row, column=4, value='Côté')
    ws.cell(row=row, column=5, value='Joueur 1')
    ws.cell(row=row, column=6, value='Joueur 2')

    for col in range(1, 7):
        c = ws.cell(row=row, column=col)
        c.font = bold_font
        c.alignment = center
        c.border = border
        c.fill = fill_header

    row += 1

    # Données
    for period in periods:
        period_name = period.get('name', '')
        period_time = period.get('time', '')
        courts = period.get('courts', [])

        for court in courts:
            court_number = court.get('number', '')

            # Côté A
            ws.cell(row=row, column=1, value=period_name)
            ws.cell(row=row, column=2, value=period_time)
            ws.cell(row=row, column=3, value=court_number)
            ws.cell(row=row, column=4, value='A')
            ws.cell(row=row, column=5, value=court['sideA']['player1']['fullName'])
            ws.cell(row=row, column=6, value=court['sideA']['player2']['fullName'])

            for col in range(1, 7):
                ws.cell(row=row, column=col).border = border
                ws.cell(row=row, column=col).alignment = left if col >= 5 else center

            row += 1

            # Côté B
            ws.cell(row=row, column=1, value=period_name)
            ws.cell(row=row, column=2, value=period_time)
            ws.cell(row=row, column=3, value=court_number)
            ws.cell(row=row, column=4, value='B')
            ws.cell(row=row, column=5, value=court['sideB']['player1']['fullName'])
            ws.cell(row=row, column=6, value=court['sideB']['player2']['fullName'])

            for col in range(1, 7):
                ws.cell(row=row, column=col).border = border
                ws.cell(row=row, column=col).alignment = left if col >= 5 else center

            row += 1

        # Joueurs en pause
        sitting = period.get('sitting', [])
        if sitting:
            ws.cell(row=row, column=1, value=period_name)
            ws.cell(row=row, column=2, value=period_time)
            ws.cell(row=row, column=3, value='Pause')
            ws.cell(row=row, column=4, value='')
            ws.cell(row=row, column=5, value=', '.join([p['fullName'] for p in sitting]))
            ws.cell(row=row, column=6, value='')

            for col in range(1, 7):
                ws.cell(row=row, column=col).border = border
                ws.cell(row=row, column=col).alignment = left if col >= 5 else center

            row += 1

    # Largeur colonnes
    widths = {
        'A': 18,
        'B': 12,
        'C': 10,
        'D': 8,
        'E': 28,
        'F': 28
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    safe_day = str(event.get('day', 'Jour')).replace('/', '-')
    filename = f"Cedule_de_la_journee_{safe_day}.xlsx"

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

@app.route('/api/export/<int:schedule_id>', methods=['GET'])
def export_schedule(schedule_id):
    schedule = Schedule.query.get(schedule_id)
    if not schedule:
        return jsonify({'error': 'Schedule not found'}), 404
    
    schedule_data = json.loads(schedule.schedule_data)
    
    # Create Excel workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    # Add sheets
    create_schedule_sheet(wb, schedule_data)
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"Pickleball_{schedule_data['event']['day']}.xlsx"
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

# ==================== ALGORITHM ====================

def generate_schedule_algorithm(event, all_players, drill_players, settings):
    """Main schedule generation with backtracking algorithm"""
    periods = []
    pairings = {
        'teammates': {p['id']: {} for p in all_players},
        'opponents': {p['id']: {} for p in all_players}
    }
    
    # Parse times
    start_minutes = time_to_minutes(event['startTime'])
    current_minutes = start_minutes
    
    # Drill period
    if event['drillMinutes'] > 0 and len(drill_players) >= 4:
        courts = generate_courts_with_constraints(drill_players, pairings, settings, True)
        periods.append({
            'name': '⚙️ Drill',
            'time': minutes_to_time(current_minutes),
            'isDrill': True,
            'courts': courts,
            'sitting': []
        })
        current_minutes += event['drillMinutes']
    
    # Regular periods
    end_minutes = time_to_minutes(event['endTime'])
    total_duration = end_minutes - current_minutes
    num_periods = total_duration // event['periodDuration']
    
    # Track play count
    play_count = {p['id']: 0 for p in all_players}
    
    for i in range(num_periods):
        # Sort by play count
        sorted_players = sorted(all_players, key=lambda p: (play_count[p['id']], -p['level']))
        
        max_playing = (len(all_players) // 4) * 4
        playing = sorted_players[:max_playing]
        sitting = sorted_players[max_playing:]
        
        courts = generate_courts_with_constraints(playing, pairings, settings, False)
        
        # Update play count
        for p in playing:
            play_count[p['id']] += 1
        
        periods.append({
            'name': f'Période {i + 1}',
            'time': minutes_to_time(current_minutes),
            'isDrill': False,
            'courts': courts,
            'sitting': sitting
        })
        current_minutes += event['periodDuration']
    
    return {
        'event': event,
        'players': all_players,
        'drillPlayers': drill_players,
        'periods': periods,
        'pairings': pairings,
        'timestamp': datetime.utcnow().isoformat()
    }

def generate_courts_with_constraints(players, pairings, settings, is_drill):
    """Generate courts using backtracking algorithm"""
    court_numbers = ['3', '4', '5', '6', '7', '8', '10', '11', '12']
    num_courts = len(players) // 4
    
    max_teammates = settings['maxTeammates']
    max_opponents = settings['maxOpponents']
    
    assignments = []
    
    def can_assign(p1, p2, p3, p4):
        tm12 = get_pairing_count(pairings['teammates'], p1['id'], p2['id'])
        tm34 = get_pairing_count(pairings['teammates'], p3['id'], p4['id'])
        
        if tm12 >= max_teammates or tm34 >= max_teammates:
            return False
        
        opp13 = get_pairing_count(pairings['opponents'], p1['id'], p3['id'])
        opp14 = get_pairing_count(pairings['opponents'], p1['id'], p4['id'])
        opp23 = get_pairing_count(pairings['opponents'], p2['id'], p3['id'])
        opp24 = get_pairing_count(pairings['opponents'], p2['id'], p4['id'])
        
        if opp13 >= max_opponents or opp14 >= max_opponents or \
           opp23 >= max_opponents or opp24 >= max_opponents:
            return False
        
        return True
    
    def score_match(p1, p2, p3, p4):
        tm12 = get_pairing_count(pairings['teammates'], p1['id'], p2['id'])
        tm34 = get_pairing_count(pairings['teammates'], p3['id'], p4['id'])
        opp_sum = sum([
            get_pairing_count(pairings['opponents'], p1['id'], p3['id']),
            get_pairing_count(pairings['opponents'], p1['id'], p4['id']),
            get_pairing_count(pairings['opponents'], p2['id'], p3['id']),
            get_pairing_count(pairings['opponents'], p2['id'], p4['id'])
        ])
        
        pairing_penalty = (tm12 + tm34) * 1000 + opp_sum * 100
        
        level_diff12 = abs(p1['level'] - p2['level'])
        level_diff34 = abs(p3['level'] - p4['level'])
        team_avg1 = (p1['level'] + p2['level']) / 2
        team_avg2 = (p3['level'] + p4['level']) / 2
        team_diff = abs(team_avg1 - team_avg2)
        
        level_penalty = (level_diff12 + level_diff34 + team_diff * 2) * 10
        
        return -(pairing_penalty + level_penalty)
    
    def backtrack(court_idx, remaining):
        if court_idx >= num_courts or len(remaining) < 4:
            return True
        
        n = len(remaining)
        candidates = []
        
        # Generate combinations
        for i in range(min(n - 3, 10)):
            for j in range(i + 1, min(n - 2, i + 10)):
                for k in range(j + 1, min(n - 1, j + 10)):
                    for l in range(k + 1, min(n, k + 10)):
                        p1, p2, p3, p4 = remaining[i], remaining[j], remaining[k], remaining[l]
                        
                        if can_assign(p1, p2, p3, p4):
                            candidates.append({
                                'players': [p1, p2, p3, p4],
                                'score': score_match(p1, p2, p3, p4)
                            })
        
        if not candidates:
            return False
        
        candidates.sort(key=lambda x: x['score'], reverse=True)
        
        for c in candidates[:5]:
            p1, p2, p3, p4 = c['players']
            
            assignments.append({'p1': p1, 'p2': p2, 'p3': p3, 'p4': p4, 'court': court_numbers[court_idx]})
            
            record_pairing(pairings['teammates'], p1['id'], p2['id'])
            record_pairing(pairings['teammates'], p3['id'], p4['id'])
            record_pairing(pairings['opponents'], p1['id'], p3['id'])
            record_pairing(pairings['opponents'], p1['id'], p4['id'])
            record_pairing(pairings['opponents'], p2['id'], p3['id'])
            record_pairing(pairings['opponents'], p2['id'], p4['id'])
            
            new_remaining = [p for p in remaining if p not in [p1, p2, p3, p4]]
            
            if backtrack(court_idx + 1, new_remaining):
                return True
            
            decrement_pairing(pairings['teammates'], p1['id'], p2['id'])
            decrement_pairing(pairings['teammates'], p3['id'], p4['id'])
            decrement_pairing(pairings['opponents'], p1['id'], p3['id'])
            decrement_pairing(pairings['opponents'], p1['id'], p4['id'])
            decrement_pairing(pairings['opponents'], p2['id'], p3['id'])
            decrement_pairing(pairings['opponents'], p2['id'], p4['id'])
            
            assignments.pop()
        
        return False
    
    backtrack(0, players)
    
    courts = []
    for a in assignments:
        courts.append({
            'number': a['court'],
            'sideA': {'player1': a['p1'], 'player2': a['p2']},
            'sideB': {'player1': a['p3'], 'player2': a['p4']}
        })
    
    return courts

def get_pairing_count(pairing_obj, id1, id2):
    return pairing_obj.get(id1, {}).get(id2, 0)

def record_pairing(pairing_obj, id1, id2):
    if id1 not in pairing_obj:
        pairing_obj[id1] = {}
    if id2 not in pairing_obj:
        pairing_obj[id2] = {}
    
    pairing_obj[id1][id2] = pairing_obj[id1].get(id2, 0) + 1
    pairing_obj[id2][id1] = pairing_obj[id2].get(id1, 0) + 1

def decrement_pairing(pairing_obj, id1, id2):
    if id1 in pairing_obj and id2 in pairing_obj[id1]:
        pairing_obj[id1][id2] -= 1
        if pairing_obj[id1][id2] <= 0:
            del pairing_obj[id1][id2]
    
    if id2 in pairing_obj and id1 in pairing_obj[id2]:
        pairing_obj[id2][id1] -= 1
        if pairing_obj[id2][id1] <= 0:
            del pairing_obj[id2][id1]

def time_to_minutes(time_str):
    parts = time_str.split(':')
    return int(parts[0]) * 60 + int(parts[1])

def minutes_to_time(minutes):
    h = minutes // 60
    m = minutes % 60
    return f"{h:02d}:{m:02d}"

def create_schedule_sheet(wb, schedule_data):
    ws = wb.create_sheet('Cédule de la journée')
    
    # Headers
    ws['A1'] = 'Terrain'
    ws['B1'] = 'Côté'
    col = 3
    
    for period in schedule_data['periods']:
        ws.cell(1, col, period['name'])
        ws.cell(1, col + 1, '')
        col += 2
    
    # Data
    row = 2
    for period in schedule_data['periods']:
        for court in period['courts']:
            # Côté A
            ws.cell(row, 1, court['number'])
            ws.cell(row, 2, 'A')
            ws.cell(row, 3, court['sideA']['player1']['fullName'])
            ws.cell(row, 4, court['sideA']['player2']['fullName'])
            row += 1
            
            # Côté B
            ws.cell(row, 1, court['number'])
            ws.cell(row, 2, 'B')
            ws.cell(row, 3, court['sideB']['player1']['fullName'])
            ws.cell(row, 4, court['sideB']['player2']['fullName'])
            row += 1

# ==================== INIT ====================
def init_db():
    with app.app_context():
        db.create_all()
        
        # Set default settings
        if not Setting.query.filter_by(key='maxTeammates').first():
            Setting.set('maxTeammates', 1)
            Setting.set('maxOpponents', 2)
            Setting.set('maxTeamLevelDiff', 0.4)
            Setting.set('maxMatchLevelDiff', 0.49)

init_db()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)