from flask import Flask, request, jsonify, send_file
from flask_cors import CORS  # Import de Flask-CORS
from tournament_generator import PetanqueTournament
from io import BytesIO
import os

app = Flask(__name__)

# Activer CORS pour autoriser uniquement les requêtes venant de ton site GitHub Pages
CORS(app, resources={r"/*": {"origins": "https://sougil.github.io"}})

@app.route('/generate_tournament', methods=['POST'])
def generate_tournament():
    data = request.get_json()
    team_type = data.get('team_type')
    num_players = data.get('num_players')
    num_matches = data.get('num_matches')

    try:
        # Initialise et génère le tournoi
        tournament = PetanqueTournament(team_type, num_players, num_matches)
        tournament.generate_matches()

        # Crée le fichier Excel en mémoire
        excel_buffer = BytesIO()
        wb = tournament.create_workbook()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        # Renvoie le fichier Excel
        filename = f"Tournoi_Petanque_{team_type.capitalize()}.xlsx"
        return send_file(excel_buffer, download_name=filename, as_attachment=True)

    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
