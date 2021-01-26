import os

home_path = os.getenv('USERPROFILE')
egrul_parser_path = home_path + '\egrul_parser'
if not os.path.exists(egrul_parser_path):
    os.mkdir(egrul_parser_path)

print(egrul_parser_path)