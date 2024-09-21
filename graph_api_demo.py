import webbrowser
import requests
import msal
from msal import PublicClientApplication

APPLICATION_ID = ''
CLIENT_SECRET  = ''
authority_url  = 'https://login.microsoftonline.com/common'
base_url       = 'https://graph.microsoft.com/v1.0/'

SCOPES = ['User.Read', 'User.Export.All', 'Mail.Read']

# method 1: Authentication with authorization code
client_instance = msal.ConfidentialClientApplication(
    client_id   = APPLICATION_ID,
    client_credential = CLIENT_SECRET,
    authority         = authority_url,
)

authorization_request_url = client_instance.get_authorization_request_url(SCOPES)
# webbrowser.open(authorization_request_url, new=True)
print(authorization_request_url)


# authorization_code = ''

authorization_code = input('Enter code')

access_token = client_instance.acquire_token_by_authorization_code(
    code = authorization_code,
    scopes = SCOPES
)

access_token_id = access_token['access_token']
headers = {'Authorization': 'Bearer ' + access_token_id}

endpoint = base_url + 'me'
response = requests.get(endpoint, headers=headers)

# Imprimir el contenido completo de la respuesta
if response.status_code == 200:
    # print(response.json())
    endpoint = base_url + 'me/mailFolders/inbox/messages?$filter=isRead eq false'
    
    responseEmail = requests.get(endpoint, headers=headers)

    # Imprimir los correos no leídos
    if responseEmail.status_code == 200:
        
        mensajes_no_leidos = responseEmail.json().get('value', [])
        
        # print(mensajes_no_leidos)
        
        for mensaje in mensajes_no_leidos:
            print(f"Subject: {mensaje['subject']}")
            print(f"Subject: {mensaje['bodyPreview']}")
            print(f"From: {mensaje['from']['emailAddress']['address']}")
            print(f"Received: {mensaje['receivedDateTime']}")
            print("-----")
            
    else:
        print(f"Error: {responseEmail.status_code} - {responseEmail.text}")
else:
    print(f"Error: {response.status_code} - {response.text}")


#       "inferenceClassification":"focused",
#       "body":{
#          "contentType":"html",
#          "content":"<html><head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><style type=\"text/css\">\r\n<!--\r\n.MLetter\r\n\t{font-family:Tahoma,Verdana,Arial,sans-serif;\r\n\tfont-size:14px;\r\n\tcolor:#454545;\r\n\tline-height:normal;\r\n\tdirection:ltr;\r\n\tpadding:0px 2px}\r\n.MLetter .MText\r\n\t{font-size:14px;\r\n\tline-height:20px}\r\n.MLetter .MSubhead\r\n\t{font-family:\\'Segoe UI Web Semibold\\',\\'Segoe UI Web Regular\\',\\'Segoe UI\\',\\'Helvetica Neue Medium\\',Arial;\r\n\tcolor:#0072C6;\r\n\tfont-size:19px;\r\n\tline-height:25px}\r\n.MLetter .MPreHeader\r\n\t{display:none}\r\n.MLetter .MPreHeader .MText\r\n\t{color:#ffffff}\r\n.MLetter .MHeader\r\n\t{background-color:#0072C6;\r\n\tcolor:#ffffff;\r\n\tpadding:50px 30px 50px}\r\n.MLetter .MHeader .MTitle\r\n\t{font-family:\\'Segoe UI Web Light\\',\\'Segoe UI Light\\',\\'Segoe UI Web Regular\\',\\'Segoe UI\\',\\'Helvetica Neue UltraLight\\',Arial;\r\n\tfont-size:40px;\r\n\tline-height:40px}\r\n.MLetter .MContent .MPara\r\n\t{padding-top:30px}\r\n.MLetter .MContent a\r\n\t{color:#0072C6;\r\n\ttext-decoration:none}\r\n.MLetter .MFooter\r\n\t{background-color:#bbbbbb;\r\n\tcolor:#ffffff;\r\n\tfont-family:Arial;\r\n\tfont-size:11px;\r\n\tpadding:30px}\r\n.MLetter .MFooter .MPara\r\n\t{padding-top:10px}\r\n.MLetter .MFooter a\r\n\t{color:#ffffff}\r\n-->\r\n</style></head><body><div class=\"MLetter\"><div class=\"MContent\" style=\"padding:22px 30px\"><div class=\"MPara\"><div class=\"MText\">Hello Andres Felipe Pizo Luligo, </div></div><div class=\"MPara\"><div class=\"MText\">To continue sending messages, please <a href=\"https://outlook.live.com/owa/\" target=\"_blank\">sign in</a> and validate your Outlook.com account. </div></div><div class=\"MPa</div></div><div class=\"MFooter\"><div>Microsoft respects your privacy. To learn more, please read our online <a href=\"http://go.microsoft.com/fwlink/p/?LinkId=253457\">Privacy Statement</a>. </div><div class=\"MPara\">Microsoft Corporation, One Microsoft Way, Redmond, WA 98052-6399, USA © 2021 Microsoft Corporation. All rights reserved. </div></div></div></body></html>"
#       },
#       "sender":{
#          "emailAddress":{
#             "name":"Outlook.com Team",
#             "address":"member_services@outlook.com"
#          }
#       },
#       "from":{
#          "emailAddress":{
#             "name":"Outlook.com Team",
#             "address":"member_services@outlook.com"
#          }
#       },
#       "toRecipients":[
#          {
#             "emailAddress":{
#                "name":"Andres Pizo",
#                "address":""
#             }
#          }
#       ],
#       "ccRecipients":[
         
#       ],
#       "bccRecipients":[
         
#       ],
#       "replyTo":[
         
#       ],
#       "flag":{
#          "flagStatus":"notFlagged"
#       }
#    }
# ]