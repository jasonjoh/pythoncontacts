# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from urllib.parse import quote
import requests
import json
import base64
import logging
from contacts.clientreg import client_registration

# Constant strings for OAuth2 flow
# The authorize URL that initiates the OAuth2 authorization code grant flow for user consent
authorize_url = 'https://login.windows.net/common/oauth2/authorize?client_id={0}&redirect_uri={1}&response_type=code'
# The token endpoint, where the app sends the auth code to get an access token
access_token_url = 'https://login.windows.net/common/oauth2/token'

# The discovery service resource and endpoint are constant
discovery_resource = 'https://api.office.com/discovery/'
discovery_endpoint = 'https://api.office.com/discovery/v1.0/me/services'

# Plugs in client ID and redirect URL to the authorize URL
# App will call this to get a URL to redirect the user for sign in
def get_authorization_url(redirect_uri):
    logger = logging.getLogger('django')
    logger.debug('Entering get_authorization_url.')
    logger.debug('  redirect_uri: {0}'.format(redirect_uri))
    
    authorization_url = authorize_url.format(client_registration.client_id(), quote(redirect_uri))
    
    logger.debug('Authorization url: {0}'.format(authorization_url))
    logger.debug('Leaving get_authorization_url.')
    return authorization_url

# Once the app has obtained an authorization code, it will call this function
# The function will request an access token for the discovery service, then
# call the discovery service to find resource IDs and endpoints for all services
# the app has permssions for
def get_access_info_from_authcode(auth_code, redirect_uri):
    logger = logging.getLogger('django')
    logger.debug('Entering get_access_info_from_authcode.')
    logger.debug('  auth_code: {0}'.format(auth_code))
    logger.debug('  redirect_uri: {0}'.format(redirect_uri))
    
    logger.debug('Sending request to access token endpoint.')
    post_data = { 'grant_type' : 'authorization_code',
                  'code' : auth_code,
                  'redirect_uri' : redirect_uri,
                  'resource' : discovery_resource,
                  'client_id' : client_registration.client_id(),
                  'client_secret' : client_registration.client_secret() }
    r = requests.post(access_token_url, data = post_data)
    logger.debug('Received response from token endpoint.')
    logger.debug(r.json())
    
    # Get the discovery service access token and do discovery
    try:
        discovery_service_token = r.json()['access_token']
        logger.debug('Extracted access token from response: {0}'.format(discovery_service_token))
    except:
        logger.debug('Exception encountered, setting token to None.')
        discovery_service_token = None
        
    if (discovery_service_token):
        # Add the refresh token to the dictionary to be returned
        # so that the app can use it to request additional access tokens
        # for other resources without having to re-prompt the user.
        discovery_result = do_discovery(discovery_service_token)
        logger.debug('Discovery completed.')
        discovery_result['refresh_token'] = r.json()['refresh_token']
        
        # Get the user's email from the access token and add to the
        # dictionary to be returned.
        json_token = parse_token(discovery_service_token)
        logger.debug('Discovery token after parsing: {0}'.format(json_token))
        discovery_result['user_email'] = json_token['upn']
        logger.debug('Extracted email from token: {0}'.format(json_token['upn']))
        logger.debug('Leaving get_access_info_from_authcode.')
        return discovery_result
    else:
        logger.debug('Leaving get_access_info_from_authcode.')
        return None

# This function calls the discovery service and parses
# the result. It builds a dictionary of resource IDs and API endpoints
# from the results.
def do_discovery(token):
    logger = logging.getLogger('django')
    logger.debug('Entering do_discovery.')
    logger.debug('  token: {0}'.format(token))
    
    headers = { 'Authorization' : 'Bearer {0}'.format(token),
                'Accept' : 'application/json' }
    r = requests.get(discovery_endpoint, headers = headers)
    
    discovery_result = {}
    
    for entry in r.json()['value']:
        capability = entry['capability']
        logger.debug('Capability found: {0}'.format(capability))
        discovery_result['{0}_resource_id'.format(capability)] = entry['serviceResourceId']
        discovery_result['{0}_api_endpoint'.format(capability)] = entry['serviceEndpointUri']
        logger.debug('  Resource ID: {0}'.format(entry['serviceResourceId']))
        logger.debug('  API endpoint: {0}'.format(entry['serviceEndpointUri']))
        
    logger.debug('Leaving do_discovery.')
    return discovery_result
    
# Once the app has obtained access information (resource IDs and API endpoints)
# it will call this function to get an access token for a specific resource. 
def get_access_token_from_refresh_token(refresh_token, resource_id):
    logger = logging.getLogger('django')
    logger.debug('Entering get_access_token_from_refresh_token.')
    logger.debug('  refresh_token: {0}'.format(refresh_token))
    logger.debug('  resource_id: {0}'.format(resource_id))
    
    post_data = { 'grant_type' : 'refresh_token',
                  'client_id' : client_registration.client_id(),
                  'client_secret' : client_registration.client_secret(),
                  'refresh_token' : refresh_token,
                  'resource' : resource_id }
                  
    r = requests.post(access_token_url, data = post_data)
    
    logger.debug('Response: {0}'.format(r.json()))
    # Return the token as a JSON object
    logger.debug('Leaving get_access_token_from_refresh_token.')
    return r.json()
    
# This function takes the base64-encoded token value and breaks
# it into header and payload, base64-decodes the payload, then
# loads that into a JSON object.
def parse_token(encoded_token):
    logger = logging.getLogger('django')
    logger.debug('Entering parse_token.')
    logger.debug('  encoded_token: {0}'.format(encoded_token))

    try:
        # First split the token into header and payload
        token_parts = encoded_token.split('.')
        
        # Header is token_parts[0]
        # Payload is token_parts[1]
        logger.debug('Token part to decode: {0}'.format(token_parts[1]))
        
        decoded_token = decode_token_part(token_parts[1])
        logger.debug('Decoded token part: {0}'.format(decoded_token))
        logger.debug('Leaving parse_token.')
        return json.loads(decoded_token)
    except:
        return 'Invalid token value: {0}'.format(encoded_token)
    
def decode_token_part(base64data):
    logger = logging.getLogger('django')
    logger.debug('Entering decode_token_part.')
    logger.debug('  base64data: {0}'.format(base64data))

    # base64 strings should have a length divisible by 4
    # If this one doesn't, add the '=' padding to fix it
    leftovers = len(base64data) % 4
    logger.debug('String length % 4 = {0}'.format(leftovers))
    if leftovers == 2:
        base64data += '=='
    elif leftovers == 3:
        base64data += '='
    
    logger.debug('String with padding added: {0}'.format(base64data))
    decoded = base64.b64decode(base64data)
    logger.debug('Decoded string: {0}'.format(decoded))
    logger.debug('Leaving decode_token_part.')
    return decoded.decode('utf-8')
        
        
        
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