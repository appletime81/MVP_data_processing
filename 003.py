import requests



rep = requests.get('https://jsonmock.hackerrank.com/api/iot_devices/search?status=RUNNING')
print(type(rep.json()))