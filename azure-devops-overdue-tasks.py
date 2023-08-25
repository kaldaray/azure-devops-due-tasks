import requests
import base64
import os
import logging
import sys
import json
import re
import jinja2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from vsts.vss_connection import VssConnection
from msrest.authentication import BasicAuthentication
from vsts.work_item_tracking.v4_1.models.wiql import Wiql
from datetime import datetime
from time import strftime, gmtime
from collections import defaultdict

def enable_logging():
    logFileName = os.path.basename(sys.argv[0]).split('.')[0]
    logDirName = os.path.dirname(os.path.realpath(__file__))
    logFilePath = logDirName + '/logs/' + logFileName + '.log'
    logging.basicConfig(
    format='%(asctime)s %(levelname)s %(message)s',
    level=logging.INFO,
    handlers=[
        logging.FileHandler(logFilePath, mode='a'),
        logging.StreamHandler()
        ]
    )

def find_email(string):
    match = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', string)
    return match.group(0)

def get_fullname(data):
    splitData = data.split()
    fullname = splitData[0] + " " + splitData[1]
    return fullname

def structuring_list(list_item,org):
    dictonaryOfItems = [
            list_item.fields["System.AssignedTo"],            
            list_item.id,
            list_item.fields["System.WorkItemType"],
            list_item.fields["System.Title"],
            list_item.fields["Microsoft.VSTS.Scheduling.DueDate"],
            list_item.fields["System.TeamProject"],
            "https://"+org+".visualstudio.com/"+str(list_item.fields["System.TeamProject"])+"/_workitems/edit/"+str(list_item.id)]
    return dictonaryOfItems

def send_mail(configuration, email_to, data):
    try:
        message = MIMEMultipart()
        message['Subject'] = 'List of due tasks'
        message['From'] = str(configuration['email_from'])
        message['To'] = email_to

        current_date = datetime.now()
        current_date = current_date.strftime("%d-%m-%Y")

        templateOfEmail = '''
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>List of due tasks</title>
    <style>
        .table {
        width: 100%;
        margin-bottom: 20px;
        border: 1px solid #dddddd;
        border-collapse: collapse; 
        }
        .table th {
            font-weight: bold;
            padding: 5px;
            background: #efefef;
            border: 1px solid #dddddd;
        }
        .table td {
            border: 1px solid #dddddd;
            padding: 5px;
        }
    </style>
  </head>
  <body>
    <div id="content">
<p>List of your due tasks by {{current_date}}!</p>
<table class="table">
    <thead>
    <tr>
        <th scope="col" style="width: 15%;">Project</th>
        <th scope="col" style="width: 20%;">Task</th>
        <th scope="col" style="width: 15%;">Deadline</th>
        <th scope="col" style="width: 50%;">URL</th>
    </tr>
    </thead>
	<tbody>
{% for item in data %}
		<tr>
			<td>{{ item[4] }}</td>
			<td>{{ item[2] }}</td>
			<td>{{ item[3][:-10] }}</td>
			<td><a href="{{ item[5] }}">{{ item[5] }}</a></td>
		</tr>
{% endfor %}
	</tbody>
</table>
    </div>   
  </body>
</html>
'''
        content = {}
        content['current_date'] = current_date
        content['data'] = data
        content['email_to'] = email_to

        emailMessage = jinja2.Template(templateOfEmail, trim_blocks=True).render(content)

        message.attach(MIMEText(emailMessage, 'html'))
        
        server = smtplib.SMTP(configuration['email_host'])
        server.ehlo()
        server.starttls()
        server.login(configuration['email_from'], configuration['email_password'])
        server.sendmail(str(configuration['email_from']), email_to, message.as_string())
        logging.info("The email was successfully sent to %s" % email_to)
        server.quit()

    except Exception as e:
        logging.error("Error of send email %s" % e)
        sys.exit()


if __name__ == "__main__":
    enable_logging()
    currentDir = os.path.dirname(__file__)
    configFile = os.path.join(currentDir, 'conf/azure_devops_connector.json')
    listOfOverdueTasks=[]

    try:
        with open(configFile) as configJson:
            logging.info("Read config file")
            configuration = json.load(configJson)
            pat = configuration['pat_token']
            org = configuration['org_name']
            authorization = str(base64.b64encode(bytes(':'+pat, 'ascii')), 'ascii')

            headers = {
                'Accept': 'application/json',
                'Authorization': 'Basic '+authorization
            }

            credentials = BasicAuthentication('', pat)
            connection = VssConnection(base_url="https://dev.azure.com/"+org, creds=credentials)

            stringQuery = """Select [System.AssignedTo], [System.Id], [System.Title], [System.State], [System.WorkItemType], [Microsoft.VSTS.Scheduling.DueDate], [System.TeamProject]
                    From WorkItems 
                    WHERE
                        [System.WorkItemType] IN ('User Story', 'Task', 'Feature')
                        AND [System.State] IN ('Active','New','In Test') AND [System.AssignedTo] !='' and [Microsoft.VSTS.Scheduling.DueDate] !='' ORDER BY [System.AssignedTo]"""
            wiql = Wiql(
                    query=stringQuery
                )
            
            wit_client = connection.get_client('vsts.work_item_tracking.v4_1.work_item_tracking_client.WorkItemTrackingClient')
            wiql_results = wit_client.query_by_wiql(wiql).work_items
            if wiql_results:
                    # Get WorkItem по id
                    work_items = (
                        wit_client.get_work_item(int(res.id)) for res in wiql_results
                    )
            
            listAllTasks=[]
            for work_item in work_items:
                # Save parameters [System.AssignedTo], [System.Id], [System.Title], [System.State], [System.WorkItemType], [Microsoft.VSTS.Scheduling.DueDate]
                listAllTasks.append(structuring_list(work_item,org))

            today = datetime.now()
            date_format = "%Y-%m-%d"
            todayIsoFormat=today.strftime(date_format)

            overDueDateTasksEmployee = defaultdict(list)
            i =0
            logging.info("Get of overdue tasks by employees")
        
            for item in listAllTasks:
                dueDate=datetime.strptime(item[4][:-10],date_format)
                newDate=dueDate.strftime(date_format)
                if newDate < todayIsoFormat:
                     overDueDateTasksEmployee[item[0]].append(item[1:7])
            logging.info("Sending emails to employees")
            for key,overdueTasks in overDueDateTasksEmployee.items():
               emailEmployee=find_email(key)
               send_mail(configuration,emailEmployee,overdueTasks)

    except Exception as e:
        logging.error("Error of parse config file %s" %e)
        sys.exit()