# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.shortcuts import render
from django.http import HttpResponseRedirect, HttpResponse
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.views import generic
from django.core.urlresolvers import reverse
from django.core.exceptions import ObjectDoesNotExist
from contacts.models import Office365Connection
import contacts.o365service
import traceback

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
        
        # For now just return the token and the user's email, the page will display it.
        context = { 'token': connection_info.access_token, 'user_email': connection_info.user_email }
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