import time

from ark_rcon.client import Client

returners_tribe_id = "1360815053"

while True:
    with Client('192.96.204.177', 11030, timeout=25) as client:
        client.login('DbJ97thHt%T8gDdH%c_HVIE13')
        # response = client.say('Hello World')
        # client.chat('Hello World')
        # response = client.run("messageoftheday")
        # response = client.players
        # response = client.getGameLog()
        # response = client.run('getchat')
        response = client.run('getgamelog')

        response_players = client.run("listplayers")
        # response = client.players
        # for player in response:
        #     client.run(f'serverchattoplayer "{player}" "Hey there, {player}."')
        #     print(player)
        print(response_players)
        # namelist = response_players.split("\n")
        #
        # namedict = [
        #     {'name': name.strip(), 'id': player_id.strip()}
        #     for line in namelist if line.strip()
        #     for name, player_id in [line.split('. ', 1)[1].split(',', 1)]
        # ]

        # for player in namedict:
        #     player_id = player['id']
        #     player_name = player['name']
        #     client.run(f'serverchattoplayer "{player_name}" "Hey there, {player_name}. A gift for you."')
        #     string_cheat = '/Game/PrimalEarth/CoreBlueprints/Resources/PrimalItemResource_Fibers.PrimalItemResource_Fibers\\'
        #     print("The id is:", player_id)
        #     cmd = (
        #         'GiveItemToPlayer 806861099 '
        #         '"Blueprint\'/Game/PrimalEarth/CoreBlueprints/Resources/PrimalItemResource_Fibers.'
        #         'PrimalItemResource_Fibers\'" 1 1 false'
        #     )
        #     client.run(cmd)
        print(client.run(f'GetTribeIdPlayerList "{returners_tribe_id}"'))
        print(namedict)
        print(response)
        time.sleep(30)