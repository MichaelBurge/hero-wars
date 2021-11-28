#!/usr/bin/env nix-shell
#! nix-shell -i python3 -p python3 pythonPackages.XlsxWriter

import json 
import itertools
import xlsxwriter
from datetime import datetime
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('asgard_file', type=str, help='file containing JSON to convert to spreadsheet')
parser.add_argument('guild_file', type=str, help='file containing guild data JSON')
parser.add_argument('heroes_file', type=str, help='file containing hero and pet data')

ALL_COLORS = [
    "NONE",
    "W",
    "G","G+1",
    "B","B+1","B+2",
    "V","V+1","V+2","V+3",
    "O","O+1","O+2","O+3","O+4",
    "R","R+1","R+2"
]

ALL_PETS = {

}

# Output 2: 

# Response 0: Not used
# Response 1: Summary of damage
# Response 2: Minions - not used

def lookup_player(players_data, player_id):
    """
    {
        results: {
            1: {
                result: {
                    response: {
                        clan: {
                            members: {
                                player_id: {
                                    name:,
                                },
                                ...
                            }
                        }
                    }
                }
            }
        }
    }
    """
    try:
        return players_data["results"][1]["result"]["response"]["clan"]["members"][player_id]["name"]
    except KeyError as e:
        print("Unknown player: " + str(e))
        return None

def lookup_hero(hero_data, hero_id):
    """
    Input:
    {
        "heroes": {
            "1": {
                name: "Astaroth"
            }
        },
        "pets": {
            "2": {
                name: "Albus
            }
        }
    }
    """
    hero  = hero_data["heroes"][hero_id]
    return hero

def lookup_pet(hero_data, pet_id):
    """
    Input:
    {
        "heros": {
            "1": {
                name: "Astaroth"
            }
        },
        "pets": {
            "2": {
                name: "Albus
            }
        }
    }
    """
    if pet_id == 0:
        return None
    PET_ID_START = 6000
    pet_idx = pet_id - PET_ID_START
    return hero_data["pets"][pet_idx]

def lookup_week(json):
    return json["results"][1]

def lookup_color(color):
    if color < len(ALL_COLORS):
        return ALL_COLORS[color]
    else:
        raise Exception("Unknown color id " + str(color))

def add_damage_summaries_page(workbook, summary_data, guild_data):
    """
    Output: Player|BossDamage|Boss Attacks|Morale Points|Minion Attacks
    Input:
    {
        result: {
            response: {
                $PLAYER_ID: {
                    bossDamage:,
                    nodesPoints:,
                    nodesAttemptsSpent:,
                    bossAttemptsSpent:
                }
            }
        }
    }
    """
    worksheet = workbook.add_worksheet("Player Summaries")
    MAX_BOSS_ATTEMPTS = 5
    MAX_MINION_ATTEMPTS = 9
    format_error = workbook.add_format({'bold': True, 'bg_color': 'red', 'font_color': 'yellow'})
    worksheet.write(0,0,"Player")
    worksheet.write(0,1,"Boss Damage")
    worksheet.write(0,2,"Boss Attempts")
    worksheet.write(0,3,"Minions Points")
    worksheet.write(0,4,"Minions Attempts")
    sorted_stats = sorted(summary_data["result"]["response"].items(), key=lambda x: int(x[1]["bossDamage"]), reverse=True)
    for row_id, (player_id, player_stats) in zip(itertools.count(1), sorted_stats):
        player_name = lookup_player(guild_data, player_id)
        if player_name is None:
            worksheet.write(row_id, 0, "Unknown player: " + player_id, format_error)
        else:
            worksheet.write(row_id, 0, player_name)
        worksheet.write(row_id, 1, player_stats["bossDamage"])
        bossAttemptsSpent = player_stats["bossAttemptsSpent"]
        worksheet.write(row_id, 2, bossAttemptsSpent, None if bossAttemptsSpent == MAX_BOSS_ATTEMPTS else format_error)
        worksheet.write(row_id, 3, player_stats["nodesPoints"])
        nodesAttemptsSpent = player_stats["nodesAttemptsSpent"]
        worksheet.write(row_id, 4, nodesAttemptsSpent, None if nodesAttemptsSpent == MAX_MINION_ATTEMPTS else format_error)

def add_match_detail_page(workbook, summary_data, guild_data, hero_data):
    """
    Output:
    Pet schema = Name|Color|Power
    Hero schema = Name|Color|Power|HP|Magic Penetration|Armor Penetration|Patroned Pet's Name|Patron Pet's Patronage Power
    Main schema = Datetime|Boss Ending Level|# of Bosses Fought|Total Damage to Boss|Damage to Boss #1|Damage to Boss #2|Hero1|Hero2|Hero3|Hero4|Hero5|Main Pet
    Input:
    {
        result:
        {
            response: {
                $PLAYER_ID: {
                    $MATCH_ID: {
                        attackers: {
                            $HERO_ID: { power, color, hp, magicPenetration, armorPenetration, favorPetId, favorPower }
                        },
                        defenders: { "1": { level, ...}},
                        effects: [ $effectName: number ],
                        startTime: "1637585752",
                        result: {
                            damage: { "1": $CURRENT_BOSS, "2": $NEXT_BOSS}],
                            progress: [
                                {
                                    defenders: {
                                        heroes: {
                                            hero_id: {
                                                extra: {
                                                    damageTaken:,
                                                    damageTakenNextLevel
                                                }
                                            }
                                        }
                                    }
                                }
                            ]
                            }
                    },
                    ...
    }}}}
    """
    worksheet = workbook.add_worksheet("Boss Match Detail")
    format_error = workbook.add_format({'bold': True, 'bg_color': 'red', 'font_color': 'yellow'})
    format_warning = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'font_color': 'black'})
    format_integer = workbook.add_format({'num_format': 1})

    pos = [ 0, 0 ]
    def write_column(x, format=None, canError=True):
        worksheet.write(pos[0], pos[1], x, format_error if canError and format is None and x is None else format)
        pos[1] += 1
    def finish_row():
        pos[0] += 1
        pos[1] = 0
    def write_hero(hero):
        hero_name = lookup_hero(hero_data, hero["id"])
        write_column(hero_name, canError=True)
        write_column(lookup_color(hero["color"]))
        write_column(hero["power"], format_integer)
        write_column(hero["hp"], format_integer)
        write_column(hero.get("magicPenetration", 0), format=format_integer) # This can be non-existent if zero.
        write_column(hero.get("armorPenetration", 0), format=format_integer)
        write_column(lookup_pet(hero_data, hero["favorPetId"]), canError=True)
        write_column(hero["favorPower"], format_integer)

    def write_pet(pet):
        if pet is None:
            write_column(None, format_warning)
            write_column(None, format_warning)
            write_column(None, format_warning)
        else:
            pet_name = lookup_pet(hero_data, pet["id"])
            write_column(pet_name, format_error if pet_name is None else None)
            write_column(lookup_color(pet["color"]))
            write_column(pet["power"], format_integer)


    matches = []
    pet_columns = [
        "Pet Name",
        "Pet Color",
        "Pet Power"
    ]
    hero_columns = [
        "Hero Name",
        "Hero Color",
        "Hero Power",
        "Hero HP",
        "Hero Magic Penetration",
        "Hero Armor Penetration",
        "Patroned Pet Name",
        "Pet Patron Power"
    ]
    columns = [
        "Player name",
        "Datetime",
        "Boss Level at End of Fight",
        "Number of Bosses Fought",
        "Total Damage to Boss",
        "Damage to Boss #1",
        "Damage to Boss #2"
    ] + (hero_columns * 5) + pet_columns
    for column_name in columns:
        write_column(column_name)
    finish_row()
    all_matches = []
    for player_id, matches in summary_data["result"]["response"].items():
        for match_id, match in matches.items():
            all_matches.append(((player_id, match_id), match))
    get_match_damages = lambda match: list(map(int, match["result"]["damage"].values()))
    sorted_matches = sorted(all_matches, key=lambda kv: sum(get_match_damages(kv[1])), reverse=True)
    for (player_id, match_id), match in sorted_matches:
        write_column(lookup_player(guild_data, player_id))
        write_column(datetime.utcfromtimestamp(int(match["startTime"])).isoformat())
        write_column(match["result"]["level"])
        bossProgress = list(match["progress"][0]["defenders"]["heroes"].values())[0]["extra"]
        bossProgresses = [int(bossProgress[key]) for key in [ "damageTaken", "damageTakenNextLevel"]]
        write_column(len(list(filter(lambda x: x > 0, bossProgresses))))
        write_column(sum(bossProgresses), format_integer)
        write_column(bossProgresses[0], format_integer)
        write_column(bossProgresses[1], format_integer)
        # match_damages = get_match_damages(match)
        # write_column(sum(match_damages))
        # write_column(match_damages[0])
        # write_column(match_damages[1])
        def get_attacker(hero_id):
            attackers = list(match["attackers"].values())
            if hero_id < len(attackers):
                return attackers[hero_id]
            else:
                return None
        for hero_id in range(5):
            hero = get_attacker(hero_id)
            if hero["type"] == "hero":
                write_hero(hero)
            else:
                raise Exception("Unknown hero type: " + hero["type"])
        #print(get_attacker(5))
        write_pet(get_attacker(5))
        finish_row()

def convert_json_to_xlsx(asgard_data, guild_data, hero_data):
    timestamp = asgard_data["date"] # 1638051586
    _, summary_data, minion_matches, boss_matches = asgard_data["results"]
    workbook = xlsxwriter.Workbook(datetime.utcfromtimestamp(timestamp).strftime('Asgard-%Y-%m-%dT%H:%M:%S.xlsx'))
    add_damage_summaries_page(workbook, summary_data, guild_data)
    add_match_detail_page(workbook, boss_matches, guild_data, hero_data)
    workbook.close()


def main():
    args = parser.parse_args()
    f = open(args.asgard_file)
    asgard_data = json.load(f)
    f = open(args.guild_file)
    guild_data = json.load(f)
    f = open(args.heroes_file)
    hero_data = json.load(f)

    convert_json_to_xlsx(asgard_data, guild_data, hero_data)

main()