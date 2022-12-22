from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import time, json
from config import *

caps = DesiredCapabilities.CHROME
caps['goog:loggingPrefs'] = {'performance': 'ALL'}
browser = webdriver.Chrome(path_to_driver, desired_capabilities=caps)
options = Options()
options.headless = True

browser.get(first_url)
time.sleep(3)
browser.get(second_url)


def process_browser_log_entry(entry):
    response = json.loads(entry['message'])['message']
    return response


browser_log = browser.get_log('performance')
events = [process_browser_log_entry(entry) for entry in browser_log]
events = [event for event in events if 'Network.response' in event['method']]

for i in range(len(events)):
    try:
        xrf = events[i]['params']['response']['url'].split("xrfkey=")[1]
        print(xrf)
        break
    except:
        continue

cookss = browser.get_cookies()
_gid = ''
_ga = ''
atlassian_xsrf_token = ''
jirasdsamlssologinv2 = ''
seraph_rememberme_cookie = ''
jsessionid = ''

for cook in cookss:
    if cook['name'] == '_gid':
        _gid = cook['value']
    elif cook['name'] == '_ga':
        _ga = cook['value']
    elif cook['name'] == 'atlassian.xsrf.token':
        atlassian_xsrf_token = cook['value']
    elif cook['name'] == 'JiraSDSamlssoLoginV2':
        jirasdsamlssologinv2 = cook['value']
    elif cook['name'] == 'seraph.rememberme.cookie':
        seraph_rememberme_cookie = cook['value']
    elif cook['name'] == 'JSESSIONID':
        jsessionid = cook['value']

print(_gid)
print(_ga)
print(atlassian_xsrf_token)
print(jirasdsamlssologinv2)
print(seraph_rememberme_cookie)
print(jsessionid)