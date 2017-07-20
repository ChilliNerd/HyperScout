#
# HyperScout v1.0
#
# Alliance selections made simple.
#
# (c) 2017 Matt Gallant (Team 6153)
#

import requests, openpyxl, os, sys, subprocess
from colorama import Fore, Style, init

# Algorithm Values
wins = 10
losses = -10
teleop_gears = 40
teleop_boiler = 700
auto_gears = 10
auto_boiler = 100

# Init Colorama on Windows
if (sys.platform == "win32"):
    init(convert = True)
else:
    init()

# Set Variables
global event_set
event_set = 0

# Setup Spreadsheet
if (os.path.isfile('teams.xlsx') == True):
    wb = openpyxl.load_workbook('teams.xlsx')
    sheet = wb.active
else:
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Teams"
    sheet['A1'] = "Team Number"
    sheet['B1'] = "Wins"
    sheet['C1'] = "Losses"
    sheet['D1'] = "Teleop Gears"
    sheet['E1'] = "Teleop Balls"
    sheet['F1'] = "Auto Gears"
    sheet['G1'] = "Auto Balls"
    sheet['H1'] = "Percentage"

# Set Event Code
def event():
    global event_var
    event_var = raw_input("Enter an event codename: ")
    global event_set
    event_set = 1

# Add Team
def add():
    # Get Team Number
    team_num = raw_input("Team Number: ")

    # The Blue Alliance API
    api_url = "https://www.thebluealliance.com/api/v3/team/frc" + team_num + "/event/" + event_var + "/status"
    header_key = {"X-TBA-Auth-Key" : "POlvVZJ0VUmUty1Z61ZlhcSWsqnwWcdRyI2Q7svEn3BYsP9Veojww4kLEGEQpyx9"}

    # Process Blue Alliance Data
    response = requests.get(api_url, headers=header_key).json()

    # Blue Alliance Data
    amount_wins = response['qual']['ranking']['record']['wins']
    amount_losses = response['qual']['ranking']['record']['losses']

    # User Input
    amount_teleop_gears = raw_input("# of Teleop Gears: ")
    amount_teleop_boiler = raw_input("# of Balls in Teleop Boiler Total: ")
    amount_auto_gears = raw_input("# of Auto Gears: ")
    amount_auto_boiler = raw_input("# of Balls in Auto Boiler Total: ")

    # Calculate Percentages
    w = float(amount_wins) / float(wins) * 16.66666666
    l = float(amount_losses) / float(losses) * 16.66666666
    tg = float(amount_teleop_gears) / float(teleop_gears) * 16.66666666
    tb = float(amount_teleop_boiler) / float(teleop_boiler) * 16.66666666
    ag = float(amount_auto_gears) / float(auto_gears) * 16.66666666
    ab = float(amount_auto_boiler) /  float(auto_boiler) * 16.66666666

    # Add Values
    total = w + l + tg + tb + ag + ab
    total = round(total, 2)

    # Print Results
    print("Team Percentage: " + str(total) + "%")

    # Sheet Data
    row = sheet.max_row + 1
    row = str(row)

    # Send Data to Spreadsheet
    sheet['A' + row] = int(team_num)
    sheet['B' + row] = int(amount_wins)
    sheet['C' + row] = int(amount_losses)
    sheet['D' + row] = int(amount_teleop_gears)
    sheet['E' + row] = int(amount_teleop_boiler)
    sheet['F' + row] = int(amount_auto_gears)
    sheet['G' + row] = int(amount_auto_boiler)
    sheet['H' + row] = float(total)
    wb.save('teams.xlsx')

# Info/About
def info():
    print("HyperScout v1")
    print("A faster way to scout!")
    print("(c) 2017 Matt Gallant (Team 6153)")

# View Table
def view():
    if (sys.platform == "win32"):
        os.startfile('teams.xlsx')
    elif (sys.platform == "darwin"):
        subprocess.call(['open', 'teams.xlsx'])
    elif (sys.platform == "linux2"):
        subprocess.call(['libreoffice', '--calc', 'teams.xlsx'])

# Exit
def exit_program():
    exit(0)

# While Loop
while True:
    # Determine if Event is Set
    if (event_set == 0):
        event_msg = "Needs to be set!"
    elif (event_set == 1):
        event_msg = ""
    else:
        event_msg = "Error!"

    # Start Main Screen
    print("    __  __                      _____                  __ ")
    print("   / / / /_  ______  ___  _____/ ___/_________  __  __/ /_")
    print("  / /_/ / / / / __ \/ _ \/ ___/\__ \/ ___/ __ \/ / / / __/")
    print(" / __  / /_/ / /_/ /  __/ /   ___/ / /__/ /_/ / /_/ / /_  ")
    print("/_/ /_/\__, / .___/\___/_/   /____/\___/\____/\__,_/\__/  ")
    print("      /____/_/                                          \n")

    print("(c) 2017 Matt Gallant\n")

    print("Type 'e' and press enter to setup event. " + Fore.RED + event_msg + Style.RESET_ALL)
    print("Type 'a' and press enter to add a team.")
    print("Type 'v' and press enter to view data.")
    print("Type 'i' and press enter for more info.")
    print("Type 'x' and press enter to save and exit.")

    # Get User Command
    cmd = raw_input("")

    # Determine Command
    if (cmd == "e"):
        event()
    elif (cmd == "a"):
        add()
    elif (cmd == "v"):
        view()
    elif (cmd == "i"):
        info()
    elif (cmd == "x"):
        exit_program()
    else:
        print("Command not valid!")