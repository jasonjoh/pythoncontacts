# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.db import models

# Create your models here.
# Represents a connection between a local account and an Office 365 account
class Office365Connection(models.Model):
    # The local username (the one used to sign into the website)
    username = models.CharField(max_length = 30)
    # The user's Office 365 account email address
    user_email = models.CharField(max_length = 254) #for RFC compliance
    # The access token from Azure
    access_token = models.TextField()
    # The refresh token from Azure
    refresh_token = models.TextField()
    # The resource ID for Outlook services (usually https://outlook.office365.com/)
    outlook_resource_id = models.URLField()
    # The API endpoint for Outlook services (usually https://outlook.office365.com/api/v1.0)
    outlook_api_endpoint = models.URLField()
    
    def __str__(self):
        return self.username

# Represents a contact item        
class DisplayContact:
    given_name = ''
    last_name = ''
    mobile_phone = ''
    email1_address = ''
    email1_name = ''
    email2_address = ''
    email2_name = ''
    email3_address = ''
    email3_name = ''
    id = ''
    
    # Initializes fields based on the JSON representation of a contact
    # returned by Office 365
    #   parameters:
    #     json: dict. The JSON dictionary object returned from Office 365.
    def load_json(self, json):
        self.given_name = json['GivenName']
        self.last_name = json['Surname']
        
        if (not json['MobilePhone1'] is None):
            self.mobile_phone = json['MobilePhone1']
        
        email_address_list = json['EmailAddresses']
        if (not email_address_list[0] is None):
            self.email1_address = email_address_list[0]['Address']
            self.email1_name = email_address_list[0]['Name']
        if (not email_address_list[1] is None):
            self.email2_address = email_address_list[1]['Address']
            self.email2_name = email_address_list[1]['Name']
        if (not email_address_list[2] is None):
            self.email3_address = email_address_list[2]['Address']
            self.email3_name = email_address_list[2]['Name']
        
        self.id = json['Id']
    
    # Generates a JSON payload for updating or creating a 
    # contact.
    #   parameters:
    #     return_nulls: Boolean. Controls how the EmailAddresses
    #                   array is generated. If True, empty entries
    #                   will be represented by "null". This style works
    #                   for update, and allows you to remove entries. 
    #                   If False, empty entries are skipped. This is needed
    #                   in the create scenario, because passing null for any entry
    #                   results in a 500 error.
    def get_json(self, return_nulls):
        json_string = '{'
        json_string += '"GivenName": "{0}"'.format(self.given_name)
        json_string += ',"Surname": "{0}"'.format(self.last_name)
        json_string += ',"MobilePhone1": "{0}"'.format(self.mobile_phone)
        json_string += ',"EmailAddresses": ['
        
        email_entry_added = False
        if (self.email1_address == '' and self.email1_name == ''):
            if (return_nulls == True):
                email_entry_added = True
                json_string += 'null'
        else:
            email_entry_added = True
            json_string += '{'
            json_string += '"@odata.type": "#Microsoft.OutlookServices.EmailAddress"'
            json_string += ',"Address": "{0}"'.format(self.email1_address)
            json_string += ',"Name": "{0}"'.format(self.email1_name)
            json_string += '}'
        
        if (self.email2_address == '' and self.email2_name == ''):
            if (return_nulls == True):
                if (email_entry_added == True):
                    json_string += ','
                email_entry_added = True;
                json_string += 'null'
        else:
            if (email_entry_added == True):
                json_string += ','
            email_entry_added = True;
            json_string += '{'
            json_string += '"@odata.type": "#Microsoft.OutlookServices.EmailAddress"'
            json_string += ',"Address": "{0}"'.format(self.email2_address)
            json_string += ',"Name": "{0}"'.format(self.email2_name)
            json_string += '}'
            
        if (self.email3_address == '' and self.email3_name == ''):
            if (return_nulls == True):
                if (email_entry_added == True):
                    json_string += ','
                email_entry_added = True;
                json_string += 'null'
        else:
            if (email_entry_added == True):
                json_string += ','
            email_entry_added = True;
            json_string += '{'
            json_string += '"@odata.type": "#Microsoft.OutlookServices.EmailAddress"'
            json_string += ',"Address": "{0}"'.format(self.email3_address)
            json_string += ',"Name": "{0}"'.format(self.email3_name)
            json_string += '}'
        
        json_string += ']'
        json_string += '}'
        
        return json_string
    
# MIT License: 
 
# Permission is hereby granted, free of charge, to any person obtaining 
# a copy of this software and associated documentation files (the 
# ""Software""), to deal in the Software without restriction, including 
# without limitation the rights to use, copy, modify, merge, publish, 
# distribute, sublicense, and/or sell copies of the Software, and to 
# permit persons to whom the Software is furnished to do so, subject to 
# the following conditions: 
 
# The above copyright notice and this permission notice shall be 
# included in all copies or substantial portions of the Software. 
 
# THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
# LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
# OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.