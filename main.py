import requests
import json
import jsonpath

url = "https://reqres.in/api/users?page=2"

# Get response
response = requests.get(url)
print(response)

# Get response in Json format
json_response = json.loads(response.text)
print(json_response)

# Fetch value of specific path
fetch_value = jsonpath.jsonpath(json_response, 'total_pages')
# Print value of a Specific path
print(fetch_value[0])

# Comparing the value
assert fetch_value[0] == 2