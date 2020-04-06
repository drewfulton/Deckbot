#----------------------------------------------------------------------------
# Name:        models.py
# Purpose:     Access, Manage, and Hold Data
# Author:    Drew Fulton
# Created:    April 4, 2020

import requests, json

endpoint = "https://api.trydatabook.com"
headers = {"Authorization": "Bearer "}
email = "drew@drewfulton.com"
pwd = "vojjij-wyHku8-soqmim"


class Company(object):
    ''' Object to incapsulate the overview of a company
    '''
    def __init__(self, company_name, company_id):
        self.id = company_id
        self.name = company_name
#         self.get_company_overview()
#         self.metrics = self.get_company_metrics()
    
    def get_company_overview(self):
        '''    Gets overview information from Databook API given a specfic Company object
        and returns a Company object
        '''
        path = f"/api/companies/{self.id}"
        response = get_api(f"{endpoint}{path}")
#         print(f"Request Status: {response}")
        overview = json.loads(response.content)
        for item in overview:
            setattr(self, item, overview[item])

    def get_company_metrics(self):
        '''    Gets a list of all metrics availble for a given Company object. 
        Returns List
        '''
        metric_list = []
        metric_list_plain = []
        response = get_api(f"{endpoint}/api/companies/{self.id}/metrics")
        metrics = json.loads(response.content)
        for metric in metrics:
            metric_list.append(Metric(metric, self.id))
            metric_list_plain.append(metric["name"])
#         print(f"Metric List: {metric_list_plain}")
        return metric_list
        
            
class Metric(object):
    '''    Object to hold company metrics
    '''
    def  __init__(self, metric, company_id):
        self.company_id = company_id
        for item in metric:
            setattr(self, item, metric[item])
        self.id = self._id
    
    def get_metric_details(self):
        path = f"{endpoint}/api/companies/{self.company_id}/metrics/{self.id}"
        response = get_api(path)
        details = json.loads(response.content)
        for d in details:        
            setattr(self, d, details[d])    
#         print(self.name)
    
def get_all_companies(offline=False):
    '''    Gets all companies from Databook API and returns a list of Company objects.
    Option offline is used to pull from hard coded sample list rather than from API
    '''
    all_companies = []
    if offline:
        for company in test_company_list:
            all_companies.append(Company(company[0],company[1]))
    else:
        response = get_api(f"{endpoint}/api/companies/")
        companies = json.loads(response.content)
        for company in companies:
            all_companies.append(Company(company['name'],company['_id']))
    return all_companies
    

def get_api(path):
    ''' Performs a GET from the Databook API.  If token is expired, it will renew.
    If another error occurs, error code is printed and the program is exited. 
    '''
    global headers
    response = requests.get(path, headers=headers)
    if response.status_code == 200:
        return response
    elif response.status_code == 401:
        # Refresh security token and get API again.
        headers = get_token(email, pwd)
        response = get_api(path)
        return response
    else:
        # Need to Add Error Handling for other issues
        print(f"Something went wrong.  Code: {response}")
        exit()
        

def get_token(email, pwd):
    '''    Get a security token from the Databook API
    '''
    global headers
    path = '/auth/local'
    body = {"email": email, "password": pwd}
    response = requests.post(f"{endpoint}{path}", data=body)
    content = json.loads(response.content)
    token = content["token"]
    headers = {"Authorization": f"Bearer {token}"}
    return headers

# List of Companies for Testing without a call to API
test_company_list = [
    ('Apple', '58080ff4ebbf470003ca9f7b'),
    ('Facebook','58153a9ac12a6e000f7c2bea'),
    ('Accenture','5809bcf577b61a00034adacf'),
    ('Amazon','5809bc2777b61a00034ada25'),
    ('Heineken','584d79ba083592000413cf7a'),
    ('Salesforce','5809dc87557fb000039bb28a')
    ]


