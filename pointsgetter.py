import requests
import json
import xlsxwriter
import os
from pprint import pprint

person = input('Enter GameWeek Number: ')
cwd = os.getcwd()
print(cwd)

workbook = xlsxwriter.Workbook(cwd + '/DraftTest.xlsx')
worksheet = workbook.add_worksheet()

teamdict = { 3 : 'ARS', 91 : 'BOU', 36 : 'BHA', 90 : 'BUR', 8 : 'CHE',
                31 : 'CRY', 11 : 'EVE', 38 : 'HUD', 13 : 'LEI', 14 : 'LIV',
                43 : 'MCI', 1 : 'MUN', 4 : 'NEW', 20 : 'SOU', 110 : 'STK',
                80 : 'SWA', 6 : 'TOT', 57 : 'WAT', 35 : 'WBA', 21 : 'WHU'
            }
posdict = { 1 : 'GK', 2 : 'DEF', 3 : 'MID', 4 : 'ST'}

gwnum = person

gwurl = 'https://fantasy.premierleague.com/drf/event/' + str(gwnum) + '/live'
plrurl = 'https://fantasy.premierleague.com/drf/bootstrap-static'

fpl_data = requests.get(plrurl).json()
gw_data = requests.get(gwurl).json()

row = 1
col = 0

worksheet.write(0, 0, 'Position')
worksheet.write(0, 1, 'Name')
worksheet.write(0, 2, 'Team')
worksheet.write(0, 3, 'Points')

for i, player in enumerate(fpl_data['elements']):
    firstname = fpl_data['elements'][i]['first_name']
    surname = fpl_data['elements'][i]['second_name']
    name = firstname + ' ' + surname
    positionnum = fpl_data['elements'][i]['element_type']
    teamcode = fpl_data['elements'][i]['team_code']
    teamname = teamdict[teamcode]
    position = posdict[positionnum]
    playerid = str(i + 1)
    points = gw_data['elements'][playerid]['stats']['total_points']

    print(position + " " + name + " " + teamname + " " + str(points))

    worksheet.write(row, col, position)
    worksheet.write(row, col + 1, name)
    worksheet.write(row, col + 2, teamname)
    worksheet.write(row, col + 3, points)
    row += 1

workbook.close()
