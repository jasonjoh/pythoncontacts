# Python Contacts Sample #

This sample is an ongoing project with two main goals: to show how to easily use the [Office 365 APIs](http://msdn.microsoft.com/en-us/office/office365/api/api-catalog) from Python, and to help me learn Python and Django. A couple of things to keep in mind:

- I am a complete newbie to Python and Django. Other than the "polls" app created as part of the [Django tutorial](https://docs.djangoproject.com/en/1.7/intro/tutorial01/), this is my first ever Python app. Because of that, I may do things in a "less-than-optimal" way from a Python perspective. Feel free to let me know!
- I chose to target the [Contacts API](http://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) for this sample. However, the same methodology should work for any of the REST APIs.
- I used the built in Django development server, so I haven't tested this with any production-level servers.
- I used the SQLite testing database that gets created with the Django project, so anything that gets stored database is in a local file on your development machine.

## Required software ##

- [Python 3.4.2](https://www.python.org/downloads/)
- [Django 1.7.1](https://docs.djangoproject.com/en/1.7/intro/install/)
- [Requests: HTTP for Humans](http://docs.python-requests.org/en/latest/)

## Running the sample ##

It's assumed that you have Python and Django installed before starting. Windows users should add the Python install directory and Scripts subdirectory to their PATH environment variable.

1. Download or fork the sample project.
2. Open your command prompt or shell to the directory where `manage.py` is located.
3. If you can run BAT files, run setup_project.bat. If not, run the three commands in the file manually. The last command prompts you to create a superuser, which you'll use later to logon.
4. Install the Requests: HTTP for Humans module from the command line: `pip install requests`
5. [Register the app in Azure Active Directory](https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md). The app should be registered as a web app with a Sign-on URL of "http://127.0.0.1:8000/contacts", and should be given permission to "Read users' contacts".
6. Edit the `.\contacts\clientreg.py` file. Copy the client ID for your app obtained during app registration and paste it as the value for the `id` variable. Copy the key you created during app registration  and paste it as the value for the `secret` variable. Save the file.
7. Start the development server: `python manage.py runserver`
8. You should see output like:
    Performing system checks...
    
    System check identified no issues (0 silenced).
    December 18, 2014 - 12:36:32
    Django version 1.7.1, using settings 'pythoncontacts.settings'
    Starting development server at http://127.0.0.1:8000/
    Quit the server with CTRL-BREAK.
9. Use your browser to go to http://127.0.0.1:8000/contacts.
10. Login with your superuser account.
11. You should now be prompted to connect your Office 365 account. Click the link to do so and login with an Office 365 account.
12. You should see the base64-encoded access token displayed. You can decode it at http://jwt.calebb.net/ to verify.
13. If you want to see what gets stored in the Django database for the user, go too http://127.0.0.1:8000/admin and click on the Office365 connections link. You can delete the user's record from the admin site too, in case you want to go through the consent process again.

## Release history ##

- **1.0: Initial release.** App allows user to connect an Office 365 account with a local app account. App does OAuth2 code grant flow and displays the user's access token.

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Exchange Dev Blog](http://blogs.msdn.com/b/exchangedev/)