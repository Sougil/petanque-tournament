import random
from openpyxl import Workbook

class PetanqueTournament:
    def __init__(self, team_type, num_players, num_matches):
        """
        Initialise le tournoi de pétanque.
        :param team_type: 'doublette' ou 'triplette'
        :param num_players: nombre total de joueurs
        :param num_matches: nombre de parties à jouer
        """
        self.team_type = team_type
        self.num_players = num_players
        self.num_matches = num_matches
        self.team_size = 2 if team_type == 'doublette' else 3

        if num_players % (self.team_size * 2) != 0:
            raise ValueError(f"Le nombre de joueurs doit être un multiple de {self.team_size * 2}")

        # Liste des joueurs
        self.players = list(range(1, num_players + 1))
        self.matches = []

    def generate_matches(self):
        """Génère des équipes aléatoirement pour chaque partie."""
        for _ in range(self.num_matches):
            random.shuffle(self.players)
            match_teams = [
                self.players[:self.team_size],
                self.players[self.team_size: self.team_size * 2]
            ]
            self.matches.append(match_teams)

    def create_workbook(self):
        """Crée un fichier Excel en mémoire avec les équipes générées."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Tournoi"

        # En-têtes
        ws.append(["Partie", "Equipe 1", "Equipe 2"])

        # Remplissage des matchs
        for idx, match in enumerate(self.matches, start=1):
            team1 = ", ".join(map(str, match[0]))
            team2 = ", ".join(map(str, match[1]))
            ws.append([f"Partie {idx}", team1, team2])

        return wb
