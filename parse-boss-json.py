#!/usr/bin/env nix-shell
#! nix-shell -i python3 -p python3 python3Packages.XlsxWriter python3Packages.numpy python3Packages.pyyaml

import json 
import itertools
import xlsxwriter
from datetime import datetime
import argparse
import numpy
import math
import yaml
import glob
#import torch

parser = argparse.ArgumentParser()
parser.add_argument('asgard_file', type=str, help='file containing JSON to convert to spreadsheet')
parser.add_argument('--guild_file', type=str, help='file containing guild data JSON', default='data/guild.json')
parser.add_argument('--heroes_file', type=str, help='file containing hero and pet data', default='data/heroes.yaml')
parser.add_argument('--buff_file', type=str, help='file containing buff data', default='data/asgard-buffs.yaml')
parser.add_argument('--history_format', type=str, help='file glob template to compute historical data over', default='data/asgard-*.json')

ALL_COLORS = [
    "NONE",
    "W",
    "G","G+1",
    "B","B+1","B+2",
    "V","V+1","V+2","V+3",
    "O","O+1","O+2","O+3","O+4",
    "R","R+1","R+2"
]

# Output 2: 

# Response 0: Not used
# Response 1: Summary of damage
# Response 2: Minions - not used

def all_players(players_data):
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
    for result in players_data["results"]:
        response = result["result"]["response"]
        if "clan" in response:
            members = response["clan"]["members"]
            for player_id, player in members.items():
                if player["clanRole"]:
                    yield player_id, player["name"]
        else:
            continue

def lookup_buff(buff_data, buff_id):
    """
    Input:
    {
        "buffInternalName": "buffUserName"
    }
    """
    for (buff_id_prefix, buff) in buff_data.items():
        if type(buff) == str:
            buff_name = buff
            buff_gold = 0
            buff_size = None
            buff_exact_name = False
        else:
            (buff_name, buff_gold, buff_size, buff_exact_name) = buff["name"], int(float(buff["gold"])), buff["size"], buff.get("exactName", False)
        if buff_exact_name:
            check = lambda x: x == buff_id_prefix
        else:
            check = lambda x: x.startswith(buff_id_prefix)
        if check(buff_id):
            return (buff_name, buff_gold, buff_size)
    else:
        return None

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
    clan_pred = lambda x: x["result"]
    try:
        for result in players_data["results"]:
            response = result["result"]["response"]
            if "clan" in response:
                return response["clan"]["members"][player_id]["name"]
            else:
                continue
        raise Exception("Unable to find guild member data")
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

def get_num_buffs(buff_data, buff_id, buff_amount):
    buff_name, buff_gold, buff_size = lookup_buff(buff_data, buff_id)
    if buff_size == None:
        return 1
    else:
        count = math.ceil(buff_amount / buff_size)
        assert count >= 1 and count <= 5, f"buff:{buff_name} amount:{buff_amount} size:{buff_size} count:{count}"
        return count

def get_buff_gold(buff_data, buff_id, buff_amount):
    buff_name, buff_gold, buff_size = lookup_buff(buff_data, buff_id)
    if buff_size == None:
        if buff_gold > 0:
            raise f"Buff gold is non-zero for purchasable buff: {buff_name} = {buff_gold}"
        ret = 0
    else:
        ret = buff_gold * get_num_buffs(buff_data, buff_id, buff_amount)
    #print(f"buff:{buff_name} total_amount:{buff_amount} size:{buff_size} gold:{ret}")
    return ret

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

boss_difficulties = [65, 75, 85, 95, 105, 115, 125, 130, 140, 150, 160, 160 ]

def next_difficulty(difficulty):
    return boss_difficulties[boss_difficulties.index(int(difficulty)) + 1]

def boss_damage_by_player_difficulty(match_data):
    ret = {}
    difficulties = {}
    for player_id, matches in match_data["result"]["response"].items():
        for match_id, match in matches.items():
            bossProgress = list(match["progress"][0]["defenders"]["heroes"].values())[0]["extra"]
            bossProgresses = [int(bossProgress[key]) for key in [ "damageTaken", "damageTakenNextLevel"]]
            difficulty = int(match["result"]["level"])
            ret.setdefault(player_id, {})
            for progress in bossProgresses:
                ret[player_id].setdefault(difficulty, 0.0)
                ret[player_id][difficulty] += progress
                difficulties[difficulty] = 1
                difficulty = next_difficulty(difficulty)
    return ret, sorted(difficulties.keys())

def add_damage_summaries_page(workbook, summary_data, match_data, guild_data):
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
    format_warn = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'font_color': 'black'})
    worksheet.write(0,0,"Player")
    worksheet.write(0,1,"Boss Damage")
    worksheet.write(0,2,"Boss Attempts")
    worksheet.write(0,3,"Minions Points")
    worksheet.write(0,4,"Minions Attempts")
    sorted_stats = sorted(summary_data["result"]["response"].items(), key=lambda x: int(x[1]["bossDamage"]), reverse=True)
    allplayers = all_players(guild_data)
    absent_players = [(pid, {}) for pid, name in allplayers if pid not in summary_data["result"]["response"]]
    boss_damages, difficulties = boss_damage_by_player_difficulty(match_data)
    for i, difficulty in enumerate(difficulties, 5):
        worksheet.write(0,i,f"Damage to {difficulty} boss")
    for row_id, (player_id, player_stats) in zip(itertools.count(1), sorted_stats + absent_players):
        player_name = lookup_player(guild_data, player_id)
        if player_name is None:
            worksheet.write(row_id, 0, "Unknown player: " + player_id, format_error)
        else:
            worksheet.write(row_id, 0, player_name)
        worksheet.write(row_id, 1, int(player_stats.get("bossDamage", 0)))
        bossAttemptsSpent = player_stats.get("bossAttemptsSpent", 0)
        worksheet.write(row_id, 2, int(bossAttemptsSpent), None if bossAttemptsSpent == MAX_BOSS_ATTEMPTS else format_error)
        nodesPoints = player_stats.get("nodesPoints")
        worksheet.write(row_id, 3, int(nodesPoints or 0), None if nodesPoints is not None else format_error)
        nodesAttemptsSpent = player_stats.get("nodesAttemptsSpent")
        worksheet.write(row_id, 4, int(nodesAttemptsSpent or 0), None if nodesAttemptsSpent == MAX_MINION_ATTEMPTS else format_error)
        player_boss_damages = boss_damages.get(player_id)
        for i, difficulty in enumerate(difficulties, 5):
            if player_boss_damages is None:
                worksheet.write(row_id, i, None, format_warn)
            else:
                damage = player_boss_damages.get(difficulty)
                worksheet.write(row_id, i, damage)

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
    format_percent = workbook.add_format({'num_format': 3})

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
        write_column(hero["hp"] + 40*hero["strength"], format_integer)
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
        "Replay Link",
        "Boss Level at End of Fight",
        "Number of Bosses Fought",
        "Total Damage to Boss",
        "Damage to Boss #1",
        "Damage to Boss #2",
        "Guild Morale",
        "Buffs"
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
        write_column("https://hero-wars.com?replay_id=" + match_id)
        write_column(match["result"]["level"])
        bossProgress = list(match["progress"][0]["defenders"]["heroes"].values())[0]["extra"]
        bossProgresses = [int(bossProgress[key]) for key in [ "damageTaken", "damageTakenNextLevel"]]
        write_column(len(list(filter(lambda x: x > 0, bossProgresses))))
        write_column(sum(bossProgresses), format_integer)
        write_column(bossProgresses[0], format_integer)
        write_column(bossProgresses[1], format_integer)
        write_column(match["effects"]["attackers"]["percentDamageBuff_any"], format_percent)
        buffstrings = []
        for k, v in match["effects"]["attackers"].items():
            buffstrings.append(k + ':' + str(v))
        write_column(','.join(buffstrings))
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
        write_pet(get_attacker(5))
        finish_row()

def add_buff_summary_page(workbook, match_detail, guild_data, hero_data, buff_data):
    worksheet = workbook.add_worksheet("Buff Summary")
    format_error = workbook.add_format({'bold': True, 'bg_color': 'red', 'font_color': 'yellow'})
    format_warning = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'font_color': 'black'})
    format_integer = workbook.add_format({'num_format': 1})
    format_percent = workbook.add_format({'num_format': 3})

    # First two columns are counts of each buff
    counts = {}
    for (player_id, matches) in match_detail["result"]["response"].items():
        match = list(matches.items())[0][1]
        for (buff_id, buff) in match["effects"]["attackers"].items():
            (buff_name, _, _) = lookup_buff(buff_data, buff_id)
            counts.setdefault(buff_id, 0)
            counts[buff_id] += 1
    sorted_counts = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)
    for row, (buff_id, count) in enumerate(sorted_counts):
        (buff_name, gold, size) = lookup_buff(buff_data, buff_id)
        worksheet.write(row, 0, buff_name or buff_id, format_warning if buff_name is None else None)
        worksheet.write(row, 1, count, format_integer)
    # Then is the buffs-by-player detail
    buffs_by_player = {}
    gold_by_player = {}
    for player_id, player in match_detail["result"]["response"].items():
        match = list(player.items())[0][1]
        for buff_id, buff in match["effects"]["attackers"].items():
            (buff_name, _, _) = lookup_buff(buff_data, buff_id)
            buffs_by_player.setdefault(player_id, {}).setdefault(buff_id, 0)
            buffs_by_player[player_id][buff_id] = buff
            gold_by_player.setdefault(player_id, 0)
            gold_by_player[player_id] += get_buff_gold(buff_data, buff_id, buff)
    worksheet.write(0, 4, "Player")
    worksheet.write(0, 5, "Gold Spent")
    for row, (player_id, player_buffs) in enumerate(sorted(buffs_by_player.items(), key=lambda kv: lookup_player(guild_data, kv[0])), start=1):
        worksheet.write(row, 4, lookup_player(guild_data, player_id))
        worksheet.write(row, 5, gold_by_player[player_id], format_integer)
        for col, (buff_id, buff) in enumerate(sorted(player_buffs.items(), key= lambda kv: kv[0]), start=6):
            (buff_name, gold, size) = lookup_buff(buff_data, buff_id)
            num_buffs = get_num_buffs(buff_data, buff_id, buff)
            buff_name = buff_name or buff_id
            if num_buffs == 1:
                buff_str = buff_name
            else:
                buff_str = f"{buff_name}: {num_buffs}"
            worksheet.write(row, col, buff_str)

def add_hero_summary_page(workbook, boss_matches, hero_data):
    """
    Output: Hero|Count|Estimated Damage Weight
    """

    all_matches = list(boss_matches["result"]["response"].items())
    num_heroes = len(hero_data["heroes"]) # Includes the placeholder, but whatever.
    num_matches = 0
    for _, matches in all_matches:
        for _, match in matches.items():
            num_matches += 1
    arr_damages = numpy.zeros(num_matches, dtype=float)
    arr_counts = numpy.zeros(num_heroes, dtype=int)
    arr_powers = numpy.zeros((num_matches, num_heroes), dtype=float)
    arr_presence = numpy.zeros((num_matches, num_heroes), dtype=float)
    arr_hero_team_damages = numpy.zeros((num_matches, num_heroes), dtype=float)
    arr_presence[:, 0] = 1.0 # Pet is always present
    arr_powers[:,0] = 100000 # Albus is usually present
    match_idx = 0
    for player_id, matches in all_matches:
        for match_id, match in matches.items():
            damages = list(match["progress"][0]["defenders"]["heroes"].values())[0]["extra"]
            damage1 = int(damages["damageTaken"])
            damage2 = int(damages["damageTakenNextLevel"])
            total_damage = damage1 + damage2
            arr_damages[match_idx] = total_damage
            for hero_id, attacker in match["attackers"].items():
                hero_idx = int(hero_id)
                if attacker["type"] == "hero":
                    arr_counts[hero_idx] += 1
                    arr_powers[match_idx, hero_idx] = attacker["power"] # attacker["power"]/30000.0 # math.log(1 + attacker["power"] / 30000.0)
                    arr_hero_team_damages[match_idx, hero_idx] = total_damage
                    arr_presence[match_idx, hero_idx] = 1.0
            match_idx += 1
    arr_hero_team_damages[:, 0] = arr_damages # Pet is always present
    worksheet = workbook.add_worksheet("Hero Summary")
    worksheet.write(0,0,"Hero")
    worksheet.write(0,1,"Count")
    worksheet.write(0,2,"Average power")
    worksheet.write(0,3,"Average team damage")
    worksheet.write(0,4,"Team Damage per Hero Power")
    format_integer = workbook.add_format({'num_format': 1})
    rows = []
    for hero_id in range(1, num_heroes):
        def avg_over_nonzero(arr):
            nonzeroes, = numpy.where(arr[:, hero_id] > 0)
            if len(nonzeroes) > 0:
                return numpy.mean(arr[:, hero_id], axis=0, where=arr[:, hero_id] > 0).item()
            else:
                return None
        avg_power = avg_over_nonzero(arr_powers) or 1
        avg_team_damage = avg_over_nonzero(arr_hero_team_damages) or 0
        rows.append([
            lookup_hero(hero_data, hero_id),
            arr_counts[hero_id],
            avg_power,
            avg_team_damage,
            avg_team_damage / avg_power
        ])
    rows.sort(reverse=True, key=lambda x: x[4])
    for row_id, row in enumerate(rows, start=1):
        for col_id, cell in enumerate(row):
            worksheet.write(row_id, col_id, cell, format_integer)

def add_team_summary_page(workbook, boss_matches, guild_data, hero_data):
    boss = {
        "armor": 35000,
        "meteorShowerMaxDamage": 120000
    }
    hero_role = lambda name: next([ role for role, name in hero_data["primary_roles"].items() if name in x ])
    # 1. A Candidate is a hero with 50k+ power OR is support/healer role.
    is_candidate = lambda hero: hero["power"] > 50000 or hero_role(hero) in ["support", "healer"]
    # 2. Damage type is inferred as Physical if hero has Armor Penetration(AP) stat, Magical with MP, and pure otherwise.
    def damage_stats(hero):
        if hero.contains_key("armorPenetration"):
            return hero["physicalAttack"], max(hero["armorPenetration"] - boss["armor"], 0)
        elif hero.contains_key("magicPenetration"):
            return hero["magicAttack"], max(hero["magicPenetration"] - boss["magicDefense"], 0)
        else:
            return max(hero["physicalAttack"], hero["magicAttack"]), 0
    # Checks:
    # 1. Do damage dealers have enough AP/MP for the boss?
    
    # 2. Does the team have a healer?
    # 3. Can every hero survive meteor(HP > physical attack OR armor artifact buff is present
    # 4. Is there a solution for defense orbs(Warrior, or Marksman that one-shots defense orbs and isn't Jhu, or special hero(Orion))?
    # 5. 
    #player_candidates = filter(lambda hero: hero["power"] > 50000 hero_role(hero) == )

def add_history_summary_page(workbook, history_data, guild_data):
    max_damages = {}
    difficulties = set()
    # Compute maximum damage rollup
    for asgard_data in history_data:
        one_summary, one_difficulties = boss_damage_by_player_difficulty(asgard_data)
        for difficulty in one_difficulties:
            difficulties.add(difficulty)
        for player_id, player_difficulty_damages in one_summary.items():
            for difficulty, damage in player_difficulty_damages.items():
                max_damages.setdefault(player_id, {})
                max_damages[player_id].setdefault(difficulty, damage)
                max_damages[player_id][difficulty] = max(max_damages[player_id][difficulty], damage)
    # Write summary
    worksheet = workbook.add_worksheet("Player Historical Summary")
    worksheet.write(0,0,"Player")
    for i, difficulty in enumerate(sorted(difficulties), 1):
        worksheet.write(0,i,f"Max total damage to {difficulty} boss")
    sorted_stats = sorted(max_damages.items(), key=lambda x: x[1].get(sorted(difficulties)[-2], 0.0), reverse=True)
    for row, (player_id, player_difficulty_damages) in enumerate(sorted_stats, 1):
        worksheet.write(row, 0, lookup_player(guild_data, player_id))
        for col, difficulty in enumerate(sorted(difficulties), 1):
            worksheet.write(row, col, player_difficulty_damages.get(difficulty))

def read_asgard_data_json(filename):
    f = open(filename)
    asgard_data = json.load(f)
    timestamp = asgard_data["date"] # 1638051586
    if len(asgard_data["results"]) == 4:
        _, summary_data, minion_matches, boss_matches = asgard_data["results"]
    elif len(asgard_data["results"]) == 3:
        summary_data, minion_matches, boss_matches = asgard_data["results"]
    else:
        raise Exception("Unknown number of results in JSON")
    return timestamp, summary_data, minion_matches, boss_matches

def convert_json_to_xlsx(asgard_data, guild_data, hero_data, history_data, buff_data):
    timestamp, summary_data, minion_matches, boss_matches = asgard_data
    workbook = xlsxwriter.Workbook(datetime.utcfromtimestamp(timestamp).strftime('Asgard-%Y-%m-%dT%H:%M:%S.xlsx'))
    add_damage_summaries_page(workbook, summary_data, boss_matches, guild_data)
    add_match_detail_page(workbook, boss_matches, guild_data, hero_data)
    add_buff_summary_page(workbook, boss_matches, guild_data, hero_data, buff_data)
    add_hero_summary_page(workbook, boss_matches, hero_data)
    add_team_summary_page(workbook, boss_matches, guild_data, hero_data)
    add_history_summary_page(workbook, history_data, guild_data)
    workbook.close()


def main():
    args = parser.parse_args()
    print(f"Asgard file: {args.asgard_file}")
    asgard_data = read_asgard_data_json(args.asgard_file)
    f = open(args.guild_file)
    guild_data = json.load(f)
    f = open(args.heroes_file)
    hero_data = yaml.safe_load(f)
    f = open(args.buff_file)
    buff_data = yaml.safe_load(f)

    history_data = []
    for file in glob.glob(args.history_format):
        timestamp, summary_data, minion_matches, boss_matches = read_asgard_data_json(file)
        history_data.append(boss_matches)

    convert_json_to_xlsx(asgard_data, guild_data, hero_data, history_data, buff_data)

main()