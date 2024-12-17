from flask import Flask, request, jsonify, send_file
from tournament_generator import PetanqueTournament
from io import BytesIO  # Import de BytesIO pour l'écriture en mémoire

app = Flask(__name__)

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

        # Crée le fichier Excel en mémoire avec BytesIO
        excel_buffer = BytesIO()
        wb = tournament.create_workbook()  # Méthode qui renvoie un objet Workbook
        wb.save(excel_buffer)  # Enregistre le fichier dans le buffer mémoire
        excel_buffer.seek(0)  # Replace le curseur au début du fichier en mémoire

        # Renvoie le fichier Excel directement au client
        filename = f"Tournoi_Petanque_{team_type.capitalize()}.xlsx"
        return send_file(excel_buffer, download_name=filename, as_attachment=True)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True)
