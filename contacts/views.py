# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.shortcuts import render
from django.http import HttpResponseRedirect, HttpResponse
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.views import generic
from django.core.urlresolvers import reverse
from django.core.exceptions import ObjectDoesNotExist
from contacts.models import Office365Connection, DisplayContact
import contacts.o365service
import traceback

contact_properties = '?$select=GivenName,Surname,MobilePhone1,EmailAddresses&$top=50'

# Create your views here.
# This is the index view for /contacts/
@login_required
def index(request):
    try:
        # Get the user's connection info
        connection_info = Office365Connection.objects.get(username = request.user)
        
    except ObjectDoesNotExist:
        # If there is no connection object for the user, they haven't connected their
        # Office 365 account yet. The page will ask them to connect.
        return render(request, 'contacts/index.html', None)
    
    else:
        # If we don't have an access token, request one
        # NOTE: This doesn't check if the token is expired. We could, but 
        # we'll just lazily assume it is good. If we try to use it and we
        # get an error, then we can refresh.
        if (connection_info.access_token is None or 
            connection_info.access_token == ''):
            # Use the refresh token to request a token for the Contacts API
            access_token = contacts.o365service.get_access_token_from_refresh_token(connection_info.refresh_token, 
                                                                                    connection_info.outlook_resource_id)
            
            # Save the access token
            connection_info.access_token = access_token['access_token']
            connection_info.save()
        
        user_contacts = contacts.o365service.get_contacts(connection_info.outlook_api_endpoint,
                                                          connection_info.access_token, contact_properties)
                                                          
        if (user_contacts is None):
            # Use the refresh token to request a token for the Contacts API
            access_token = contacts.o365service.get_access_token_from_refresh_token(connection_info.refresh_token, 
                                                                                    connection_info.outlook_resource_id)
            
            # Save the access token
            connection_info.access_token = access_token['access_token']
            connection_info.save()
            
            user_contacts = contacts.o365service.get_contacts(connection_info.outlook_api_endpoint,
                                                              connection_info.access_token, contact_properties)
        contact_list = list()
        
        for user_contact in user_contacts['value']:
            display_contact = DisplayContact()
            display_contact.load_json(user_contact)
            contact_list.append(display_contact)
        
        # For now just return the token and the user's email, the page will display it.
        context = { 'user_email': connection_info.user_email,
                    'user_contacts': contact_list }
        return render(request, 'contacts/index.html', context)
        
# The /contacts/connect/ action. This will redirect to the Azure OAuth
# login/consent page.
def connect(request):
    redirect_uri = 'http://127.0.0.1:8000/contacts/authorize/'
    
    url = contacts.o365service.get_authorization_url(redirect_uri)
    return HttpResponseRedirect(url)
    
# The /contacts/authorize action. This is where Azure's login/consent page
# redirects after the user consents.
def authorize(request):
    redirect_uri = 'http://127.0.0.1:8000/contacts/authorize/'
    if request.method == "GET":
        # Azure passes the auth code in the 'code' parameter
        try:
            auth_code = request.GET['code']
        except:
            return render(request, 'contacts/error.html', {
                    'error_message' : 'Connection canceled.'
                })
        else:
            # Get the user's connection info from database
            try:
                connection = Office365Connection.objects.get(username = request.user)
            except ObjectDoesNotExist:
                # If there is not one for the user, create a new one
                connection = Office365Connection(username = request.user)

            # Use the auth code to get an access token and do discovery
            access_info = contacts.o365service.get_access_info_from_authcode(auth_code, redirect_uri)
            
            if (access_info):
                try:
                    user_email = access_info['user_email']
                    refresh_token = access_info['refresh_token']
                    resource_id = access_info['Contacts_resource_id']
                    api_endpoint = access_info['Contacts_api_endpoint']
                    
                    # Save the access information in the user's connection info
                    connection.user_email = user_email
                    connection.refresh_token = refresh_token
                    connection.outlook_resource_id = resource_id
                    connection.outlook_api_endpoint = api_endpoint
                    connection.save()
                except Exception as e:
                    return render(request, 'contacts/error.html', 
                        {
                            'error_message': 'An exception occurred: {0}'.format(traceback.format_exception_only(type(e), e))
                        })
            else:
                return render(request, 'contacts/error.html', 
                    {
                        'error_message': 'Unable to connect Office 365 account.',
                    })
        
            return HttpResponseRedirect(reverse('contacts:index'))
    else:
        return HttpResponseRedirect(reverse('contacts:index'))

# The new view, used to display a blank details form for a contact.
@login_required
def new(request):
    return render(request, 'contacts/details.html', None)       
        
# The create action, invoked via POST by the details form when creating
# a new contact.
@login_required
def create(request):
    try:
        # Initialize a DisplayContact object from the posted form data
        new_contact = DisplayContact()
        new_contact.given_name = request.POST['first_name']
        new_contact.last_name = request.POST['last_name']
        new_contact.mobile_phone = request.POST['mobile_phone']
        new_contact.email1_address = request.POST['email1_address']
        new_contact.email1_name = request.POST['email1_name']
        new_contact.email2_address = request.POST['email2_address']
        new_contact.email2_name = request.POST['email2_name']
        new_contact.email3_address = request.POST['email3_address']
        new_contact.email3_name = request.POST['email3_name']
        
    except (KeyError):
        # If the form data is missing or incomplete, display an error.
        return render(request, 'contacts/error.html',
                {
                    'error_message': 'No contact data included in POST.',
                }
            )
    else:
        try:
            # Get the user's connection info
            connection_info = Office365Connection.objects.get(username = request.user)
            
        except ObjectDoesNotExist:
            # If there is no connection object for the user, they haven't connected their
            # Office 365 account yet. The page will ask them to connect.
            return render(request, 'contacts/index.html', None)
            
        else:
            result = contacts.o365service.create_contact(connection_info.outlook_api_endpoint,
                                                         connection_info.access_token,
                                                         new_contact.get_json(False))
            # Per MSDN, success should be a 201 status                                             
            if (result == 201):
                return HttpResponseRedirect(reverse('contacts:index'))
            else:
                return render(request, 'contacts/error.html',
                    {
                        'error_message': 'Unable to create contact: {0} HTTP status returned.'.format(result),
                    }
                )

# The edit view, used to display an existing contact in a details form.
# Note this view always retrieves the contact from Office 365 to get the latest version.                
@login_required
def edit(request, contact_id):        
    try:
        # Get the user's connection info
        connection_info = Office365Connection.objects.get(username = request.user)
        
    except ObjectDoesNotExist:
        # If there is no connection object for the user, they haven't connected their
        # Office 365 account yet. The page will ask them to connect.
        return render(request, 'contacts/index.html', None)
        
    else:
        if (connection_info.access_token is None or 
                connection_info.access_token == ''):
            return render(request, 'contacts/index.html', None)
            
    contact_json = contacts.o365service.get_contact_by_id(connection_info.outlook_api_endpoint,
                                                          connection_info.access_token,
                                                          contact_id, contact_properties)
                                                          
    if (not contact_json is None):
        # Load the contact into a DisplayContact object
        display_contact = DisplayContact()
        display_contact.load_json(contact_json)
        
        # Render a details form
        return render(request, 'contacts/details.html', { 'contact': display_contact })
    else:
        return render(request, 'contacts/error.html',
            {
                'error_message': 'Unable to get contact with ID: {0}'.format(contact_id),
            }
        )

# The update action, invoked via POST from the details form when editing
# an existing contact.
@login_required   
def update(request, contact_id):
    try:
        # Initialize a DisplayContact object from the posted form data
        updated_contact = DisplayContact()
        updated_contact.given_name = request.POST['first_name']
        updated_contact.last_name = request.POST['last_name']
        updated_contact.mobile_phone = request.POST['mobile_phone']
        updated_contact.email1_address = request.POST['email1_address']
        updated_contact.email1_name = request.POST['email1_name']
        updated_contact.email2_address = request.POST['email2_address']
        updated_contact.email2_name = request.POST['email2_name']
        updated_contact.email3_address = request.POST['email3_address']
        updated_contact.email3_name = request.POST['email3_name']
        
    except (KeyError):
        # If the form data is missing or incomplete, display an error.
        return render(request, 'contacts/error.html',
                {
                    'error_message': 'No contact data included in POST.',
                }
            )
    else:
        try:
            # Get the user's connection info
            connection_info = Office365Connection.objects.get(username = request.user)
            
        except ObjectDoesNotExist:
            # If there is no connection object for the user, they haven't connected their
            # Office 365 account yet. The page will ask them to connect.
            return render(request, 'contacts/index.html', None)
            
        else:
            result = contacts.o365service.update_contact(connection_info.outlook_api_endpoint,
                                                         connection_info.access_token,
                                                         contact_id,
                                                         updated_contact.get_json(True))
            
            # Per MSDN, success should be a 200 status
            if (result == 200):
                return HttpResponseRedirect(reverse('contacts:index'))
            else:
                return render(request, 'contacts/error.html',
                    {
                        'error_message': 'Unable to update contact: {0} HTTP status returned.'.format(result),
                    }
                )
        
# The delete action, invoked to delete a contact.        
@login_required
def delete(request, contact_id):
    try:
        # Get the user's connection info
        connection_info = Office365Connection.objects.get(username = request.user)
        
    except ObjectDoesNotExist:
        # If there is no connection object for the user, they haven't connected their
        # Office 365 account yet. The page will ask them to connect.
        return render(request, 'contacts/index.html', None)
        
    else:
        if (connection_info.access_token is None or 
                connection_info.access_token == ''):
            return render(request, 'contacts/index.html', None)
            
        result = contacts.o365service.delete_contact(connection_info.outlook_api_endpoint,
                                                     connection_info.access_token,
                                                     contact_id)
        
        # Per MSDN, success should be a 204 status
        if (result == 204):
            return HttpResponseRedirect(reverse('contacts:index'))
        else:
            return render(request, 'contacts/error.html',
                {
                    'error_message': 'Unable to delete contact: {0} HTTP status returned.'.format(result),
                }
            )
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