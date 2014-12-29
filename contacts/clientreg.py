# The client ID (register app in Azure AD to get this value)
id = '<PUT YOUR CLIENT ID HERE>';
# The client secret (register app in Azure AD to get this value)
secret = '<PUT YOUR CLIENT SECRET (KEY) HERE>';

class client_registration:
    @staticmethod
    def client_id():
        return id;
        
    @staticmethod
    def client_secret():
        return secret;