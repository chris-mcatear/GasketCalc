import requests

class CheckforUpdate():
    def __init__(self):
        pass
    
    def updatechecker(self):
        headers = {'Accept': 'application/vnd.github+json'}
        url = "https://api.github.com/repos/chris-mcatear/GasketCalc/releases/latest"
        response = requests.get(url, headers=headers)
        # print(response.content)
        if response.status_code == 200:
            # print("success")
            # print(response.json()["tag_name"])
            return response.json()["tag_name"]
        else:
            # print("error")
            return "Error"
    
    #print("test")