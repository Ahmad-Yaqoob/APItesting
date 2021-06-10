import json

import requests
url = "https://reqres.in/api/login"
a = {
    "email": "eve.holt@reqres.in",
    "password": "cityslicka"
     }
j = json.dumps(a)
#print(a)
print(j)