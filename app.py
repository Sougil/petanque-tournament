from flask import Flask, request, jsonify, send_file
from petanque_tournament_generator import PetanqueTournament

from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Autorise toutes les origines


@app.route('/generate_tournament', methods=['POST'])
def generate_tournament():
    data = request.get_json()
    team_type = data.get('team_type')
    num_players = data.get('num_players')
    num_matches = data.get('num_matches')

    try:
        # Crée le tournoi
        tournament = PetanqueTournament(team_type, num_players, num_matches)
        tournament.generate_matches()
        
        # Crée et envoie le fichier Excel
        filename = tournament.create_excel_file()
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True)
