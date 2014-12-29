# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.db import models

# Create your models here.
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