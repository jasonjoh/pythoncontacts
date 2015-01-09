# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
from django.conf.urls import patterns, url

from contacts import views

urlpatterns = patterns('',
    # The home view ('/contacts/')
    url(r'^$', views.index, name='index'),
    # Used to start OAuth2 flow ('/contacts/connect/')
    url(r'^connect/$', views.connect, name='connect'),
    # Used as redirect target in OAuth2 flow ('/contacts/authorize/')
    url(r'^authorize/$', views.authorize, name='authorize'),
    # Displays a form to create a new contact ('/contacts/new/')
    url(r'^new/$', views.new, name='new'),
    # Invoked to create a new contact in Office 365 ('/contacts/create/')
    url(r'^create/$', views.create, name='create'),
    # Displays an existing contact in an editable form ('/contacts/edit/<contact_id>/')
    url(r'^edit/(?P<contact_id>.+)/$', views.edit, name='edit'),
    # Invoked to update an existing contact ('/contacts/update/<contact_id>/')
    url(r'^update/(?P<contact_id>.+)/$', views.update, name='update'),
    # Invoked to delete an existing contact ('/contacts/delete/<contact_id>/')
    url(r'^delete/(?P<contact_id>.+)/$', views.delete, name='delete'),
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