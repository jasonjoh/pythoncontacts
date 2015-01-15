# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.test import TestCase
from django.core.exceptions import ObjectDoesNotExist
from contacts.models import Office365Connection
import contacts.o365service
# Create your tests here.

api_endpoint = 'https://outlook.office365.com/api/v1.0'

# TODO: Copy a valid, non-expired access token here. You can get this from
# an Office365Connection in the /admin/ page once you've successfully connected
# an account to view contacts in the app. Remember these expire every hour, so
# if you start getting 401's you need to get a new token.
access_token = ''

class MailApiTests(TestCase):
    
    def test_create_message(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        new_message_payload = '{ "Subject": "Did you see last night\'s game?", "Importance": "Low", "Body": { "ContentType": "HTML", "Content": "They were <b>awesome</b>!" }, "ToRecipients": [ { "EmailAddress": { "Address": "jasonjoh@alpineskihouse.com" } } ] }'
            
        r = contacts.o365service.create_message(api_endpoint,
                                                access_token,
                                                new_message_payload)
                                                
        self.assertEqual(r, 201, 'Create message returned {0}'.format(r))
        
    def test_get_message_by_id(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        get_messages_params = '?$top=5&$select=Subject'
    
        r = contacts.o365service.get_messages(api_endpoint,
                                              access_token,
                                              get_messages_params)
                                              
        self.assertIsNotNone(r, 'Get messages returned None.')
        
        first_message = r['value'][0]
        
        first_message_id = first_message['Id']
        
        r = contacts.o365service.get_message_by_id(api_endpoint,
                                                   access_token,
                                                   first_message_id)
                                                   
        self.assertIsNotNone(r, 'Get message by id returned None.')
        
    def test_update_message(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        get_messages_params = '?$top=5&$select=Subject'
    
        r = contacts.o365service.get_messages(api_endpoint,
                                              access_token,
                                              get_messages_params)
                                              
        self.assertIsNotNone(r, 'Get messages returned None.')
        
        first_message = r['value'][0]
        
        first_message_id = first_message['Id']
        
        update_payload = '{ "Subject" : "UPDATED" }'
        
        r = contacts.o365service.update_message(api_endpoint,
                                                access_token,
                                                first_message_id,
                                                update_payload)
                                                   
        self.assertEqual(r, 200, 'Update message returned {0}.'.format(r))
        
    def test_delete_message(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        get_messages_params = '?$top=5&$select=Subject'
    
        r = contacts.o365service.get_messages(api_endpoint,
                                              access_token,
                                              get_messages_params)
                                              
        self.assertIsNotNone(r, 'Get messages returned None.')
        
        first_message = r['value'][0]
        
        first_message_id = first_message['Id']
        
        r = contacts.o365service.delete_message(api_endpoint,
                                                access_token,
                                                first_message_id)
                                                   
        self.assertEqual(r, 204, 'Delete message returned {0}.'.format(r))
        
    def test_send_draft_message(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        # Get drafts
        get_drafts = '{0}/Me/Folders/Drafts/Messages?$select=Subject'.format(api_endpoint)
        
        r = contacts.o365service.make_api_call('GET', get_drafts, access_token)
        
        response = r.json()
        
        first_message = response['value'][0]
        
        first_message_id = first_message['Id']
        
        send_response = contacts.o365service.send_draft_message(api_endpoint,
                                                                access_token,
                                                                first_message_id)
                                                                
        self.assertEqual(r, 200, 'Send draft returned {0}.'.format(r))
        
    def test_send_new_mail(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        new_message_payload = '{ "Subject": "Sent from test_send_new_mail", "Importance": "Low", "Body": { "ContentType": "HTML", "Content": "They were <b>awesome</b>!" }, "ToRecipients": [ { "EmailAddress": { "Address": "allieb@jasonjohtest.onmicrosoft.com" } } ] }'
        
        r = contacts.o365service.send_new_message(api_endpoint,
                                                  access_token,
                                                  new_message_payload,
                                                  True)
                                                  
        self.assertEqual(r, 202, 'Send new message returned {0}.'.format(r))
        
class CalendarApiTests(TestCase):
    
    def test_create_event(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        new_event_payload = '{ "Subject": "Discuss the Calendar REST API", "Body": { "ContentType": "HTML", "Content": "I think it will meet our requirements!" }, "Start": "2015-01-15T18:00:00Z", "End": "2015-01-15T19:00:00Z", "Attendees": [ { "EmailAddress": { "Address": "alexd@alpineskihouse.com", "Name": "Alex Darrow" }, "Type": "Required" } ] }'
            
        r = contacts.o365service.create_event(api_endpoint,
                                              access_token,
                                              new_event_payload)
                                                
        self.assertEqual(r, 201, 'Create event returned {0}'.format(r))
        
    def test_get_event_by_id(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        get_events_params = '?$top=5&$select=Subject,Start,End'
    
        r = contacts.o365service.get_events(api_endpoint,
                                            access_token,
                                            get_events_params)
                                              
        self.assertIsNotNone(r, 'Get events returned None.')
        
        first_event = r['value'][0]
        
        first_event_id = first_event['Id']
        
        r = contacts.o365service.get_event_by_id(api_endpoint,
                                                 access_token,
                                                 first_event_id)
                                                   
        self.assertIsNotNone(r, 'Get event by id returned None.')
        
    def test_update_event(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        get_events_params = '?$top=5&$select=Subject,Start,End'
    
        r = contacts.o365service.get_events(api_endpoint,
                                            access_token,
                                            get_events_params)
                                              
        self.assertIsNotNone(r, 'Get events returned None.')
        
        first_event = r['value'][0]
        
        first_event_id = first_event['Id']
        
        update_payload = '{ "Subject" : "UPDATED" }'
        
        r = contacts.o365service.update_event(api_endpoint,
                                              access_token,
                                              first_event_id,
                                              update_payload)
                                                   
        self.assertEqual(r, 200, 'Update event returned {0}.'.format(r))
        
    def test_delete_event(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
            
        get_events_params = '?$top=5&$select=Subject,Start,End'
    
        r = contacts.o365service.get_events(api_endpoint,
                                            access_token,
                                            get_events_params)
                                              
        self.assertIsNotNone(r, 'Get events returned None.')
        
        first_event = r['value'][0]
        
        first_event_id = first_event['Id']
        
        r = contacts.o365service.delete_event(api_endpoint,
                                              access_token,
                                              first_event_id)
                                                   
        self.assertEqual(r, 204, 'Delete event returned {0}.'.format(r))

class ContactsApiTests(TestCase):
    
    def test_create_contact(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
        
        new_contact_payload = '{ "GivenName": "Pavel", "Surname": "Bansky", "EmailAddresses": [ { "Address": "pavelb@alpineskihouse.com", "Name": "Pavel Bansky" } ], "BusinessPhones": [ "+1 732 555 0102" ] }'
            
        r = contacts.o365service.create_contact(api_endpoint,
                                                access_token,
                                                new_contact_payload)
                                                
        self.assertEqual(r, 201, 'Create contact returned {0}'.format(r))
        
    def test_get_contact_by_id(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
        
        get_contacts_params = '?$top=5&$select=DisplayName'
    
        r = contacts.o365service.get_contacts(api_endpoint,
                                              access_token,
                                              get_contacts_params)
                                              
        self.assertIsNotNone(r, 'Get contacts returned None.')
        
        first_contact = r['value'][0]
        
        first_contact_id = first_contact['Id']
        
        r = contacts.o365service.get_contact_by_id(api_endpoint,
                                                   access_token,
                                                   first_contact_id)
                                                   
        self.assertIsNotNone(r, 'Get contact by id returned None.')
        
    def test_update_contact(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
        
        get_contacts_params = '?$top=5&$select=DisplayName'
    
        r = contacts.o365service.get_contacts(api_endpoint,
                                              access_token,
                                              get_contacts_params)
                                              
        self.assertIsNotNone(r, 'Get contacts returned None.')
        
        first_contact = r['value'][0]
        
        first_contact_id = first_contact['Id']
        
        update_payload = '{ "Surname" : "UPDATED" }'
        
        r = contacts.o365service.update_contact(api_endpoint,
                                                access_token,
                                                first_contact_id,
                                                update_payload)
                                                   
        self.assertEqual(r, 200, 'Update contact returned {0}.'.format(r))
        
    def test_delete_contact(self):
        self.assertEqual(access_token, '', 'You must copy a valid access token into the access_token variable.')
        
        get_contacts_params = '?$top=5&$select=DisplayName'
    
        r = contacts.o365service.get_contacts(api_endpoint,
                                              access_token,
                                              get_contacts_params)
                                              
        self.assertIsNotNone(r, 'Get contacts returned None.')
        
        first_contact = r['value'][0]
        
        first_contact_id = first_contact['Id']
        
        r = contacts.o365service.delete_contact(api_endpoint,
                                                access_token,
                                                first_contact_id)
                                                   
        self.assertEqual(r, 204, 'Delete contact returned {0}.'.format(r))
        
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