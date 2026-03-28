import json, base64
d = json.load(open(r'c:\Users\LocNTP\Downloads\ku_dev\teams_app\TeamsExportMess\token.json'))
t = d['token']
parts = t.split('.')
p = parts[1] + '=' * (-len(parts[1]) % 4)
payload = json.loads(base64.urlsafe_b64decode(p))
print('aud:', payload.get('aud'))
print('appid:', payload.get('appid'))
print('app:', payload.get('app_displayname'))
print()
print('ALL SCOPES:')
for s in payload.get('scp', '').split():
    print(f'  {s}')
print()
has_chat = any('chat' in s.lower() for s in payload.get('scp', '').split())
print(f'Has Chat.Read: {has_chat}')
