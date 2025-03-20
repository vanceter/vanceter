import requests
from bs4 import BeautifulSoup
import csv

# Function to fetch the tournament schedule (for demonstration purposes, let's assume it's available in HTML)
def fetch_tournament_schedule(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Assuming the games are listed in a table with class 'schedule-table'
    games = []
    for row in soup.select('.schedule-table tbody tr'):
        game = {}
        game['date'] = row.select('td.date')[0].text.strip()
        game['time'] = row.select('td.time')[0].text.strip()
        game['team1'] = row.select('td.team1')[0].text.strip()
        game['team2'] = row.select('td.team2')[0].text.strip()
        games.append(game)
    
    return games

# Function to save the tournament schedule to a CSV file
def save_to_csv(games, output_file):
    # Define CSV headers
    headers = ['Date', 'Time', 'Team 1', 'Team 2']

    with open(output_file, 'w', newline='', encoding='utf-8') as file:
        writer = csv.DictWriter(file, fieldnames=headers)
        writer.writeheader()  # Write the header row
        for game in games:
            writer.writerow(game)

# Example URL (replace with actual source)
url = 'https://www.ncaa.com/news/basketball-men/article/2025-03-18/2025-march-madness-mens-ncaa-tournament-schedule-dates'

# Fetch the schedule
games = fetch_tournament_schedule(url)

# Save the schedule to a CSV file
save_to_csv(games, 'ncaa_tournament_schedule.csv')

print("CSV file created successfully!")
