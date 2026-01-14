"""
=====================================================================================================================
Required imports for script to function as expected
=====================================================================================================================
"""
# Kaltura Library Imports
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm
import queue
import threading
from datetime import datetime
import subprocess
import os
from colorama import Fore, Back, Style
import csv
from easygui import *
import json
import requests
from urllib.parse import quote_plus
import re
import logging
import time
import configparser
import sys
import hashlib
from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph import GraphServiceClient
from azure.core.exceptions import ClientAuthenticationError, HttpResponseError
from azure.identity import ClientSecretCredential, InteractiveBrowserCredential
from KalturaClient.Plugins.Caption import *
from KalturaClient.Plugins.Metadata import *
from KalturaClient.Plugins.Core import *
from unittest import case
import asyncio
from KalturaClient import *

# Azure Library Imports

# General Python Library Imports


"""
=====================================================================================================================
Variable setup
=====================================================================================================================
"""


# Sets the proxy/prot for the script to run
# If you need to use a proxy, uncomment the following lines and set the proxy URL
# proxy = 'proxy:port'
# os.environ['http_proxy'] = proxy
# os.environ['HTTP_PROXY'] = proxy
# os.environ['https_proxy'] = proxy
# os.environ['HTTPS_PROXY'] = proxy

global sleepTimeSeconds
sleepTimeSeconds = 5

# Get the directory the script is in and set it for subprocess calls
script_dir = os.path.dirname(os.path.abspath(__file__))



"""
=====================================================================================================================
Script functions, section based on requirements and calls
=====================================================================================================================
"""

"""
=====================================================================================================================
Read configuration file for this to run
file name is config.ini
Config has general connection information for all runnable functions
=====================================================================================================================
"""


def read_config():
    # Create a ConfigParser object
    config = configparser.ConfigParser()

    # Read the configuration file
    config.read('CaptionFinder/config.ini')

    # Access values from the configuration file
    # Grab General Info from config file
    azure_url = config.get('General', 'azure_url')
    azure_audience = config.get('General', 'azure_audience')
    log_level = config.get('General', 'log_level')
    debug_mode = config.getboolean('General', 'debug')

    # Grab Azure Info from config file
    azureClientID = config.get('Azure', 'azureClientID')
    azureClientSecret = config.get('Azure', 'azureClientSecret')
    azureTenantID = config.get('Azure', 'azureTenantID')

    # Return a dictionary with the retrieved values
    config_values = {
        'azure_url': azure_url,
        'azure_audience': azure_audience,
        'log_level': log_level,
        'debug_mode': debug_mode,
        'azureClientID': azureClientID,
        'azureClientSecret': azureClientSecret,
        'azureTenantID': azureTenantID
    }

    return config_values


"""
=====================================================================================================================
End reading config function
=====================================================================================================================
"""


"""
=====================================================================================================================
Connect to Microsoft Graph API
    This will connect to the graph API's. It uses the scope set in the Azure Application
    

    The following permissions is what is set for this to run:
    - Directory.Read.All
    - Group.Read.All
    - GroupMember.Read.All
    - User.Read.All
=====================================================================================================================
"""


def connectToMSGraph():
    # Function to connect to Microsoft Graph API
    print(Fore.GREEN + "We will now connect to Microsoft Graph API")
    scopes = ['https://graph.microsoft.com/.default']
    print(Fore.GREEN + "Using the following scopes:")
    for scope in scopes:
        print(Fore.GREEN + f" - {scope}")

    # Tenant Information, pulled from the config section
    print(Fore.GREEN + "Setting the Azure credentials...")
    TENANT_ID = config_data['azureTenantID']
    CLIENT_ID = config_data['azureClientID']
    CLIENT_SECRET = config_data['azureClientSecret']

    # Initialize credential
    print(Fore.GREEN + "Initializing credentials...")
    credential = ClientSecretCredential(
        tenant_id=TENANT_ID,
        client_id=CLIENT_ID,
        client_secret=CLIENT_SECRET
    )

    print(Fore.GREEN + "Trying to Connecting to Microsoft Graph API...")
    try:
        # Initialize GraphServiceClient
        graph_client = GraphServiceClient(credential, scopes)
        print(Fore.GREEN + "Connected to Microsoft Graph API successfully!")
        print(Fore.GREEN + "Sending the client information back to the caller...")
        return graph_client
    except Exception as e:
        print(Fore.RED + f"Error connecting to Microsoft Graph API: {e}")
        return None


"""
=====================================================================================================================
End connectToMSGraph function
=====================================================================================================================
"""

"""
=====================================================================================================================
Get user email from Azure using Python method
    This method retrieves user information from Azure Active Directory using the Microsoft Graph API.
    It returns the user's display name, email, and account status.
    This function was built with a combination of the following
        MS Graph Explorer https://developer.microsoft.com/en-us/graph/graph-explorer
        Microsoft Graph Rest API, Example 3: https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=python
    This is supposed to be an async function, but it is currently synchronous and uses asyncio to handle the 
    graphClient.users.by_user_id(user_id).get(request_configuration=request_configuration) aspect. 

    There are better methods, but since this script relies on the current implementation, we will proceed with it.
=====================================================================================================================
"""


def get_User_Email_Azure_python(user_id):
    print(Fore.GREEN +
          f"Fetching user email for User ID {user_id} using Python method...")

    try:
        # Verify that the graph client is initialized
        if not graphClient:
            print(
                Fore.RED + "Graph client is not initialized. Cannot retrieve user information!")
            return {
                "displayName": "Graph Client Error",
                "mail": "Graph Client Error",
                "accountEnabled": "Graph Client Error"
            }
        else:
            print(
                Fore.GREEN + "Graph client is initialized, proceeding to fetch user information...")
    except Exception as e:
        print(
            Fore.RED + f"Error accessing graph client for User ID {user_id}: {e}")
        return {
            "displayName": "Graph Client Error",
            "mail": "Graph Client Error",
            "accountEnabled": "Graph Client Error"
        }

    # Build the request configuration to export DisplayName, Mail, and account status
    # If other fields are needed like First Name, Last Name, or the like, you'll need to add those
    # Azure uses givenName for First name and surname for Last Name
    # Example of other fields select=["displayName", "mail", "accountEnabled", "givenName", "surname"]
    query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
        select=["displayName", "mail", "accountEnabled"],
    )

    request_configuration = RequestConfiguration(
        query_parameters=query_params,
    )

    try:
        # Try to get existing loop, create new one if none exists
        try:
            loop = asyncio.get_event_loop()
            if loop.is_closed():
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)

        result = loop.run_until_complete(
            graphClient.users.by_user_id(user_id).get(
                request_configuration=request_configuration)
        )

        # Return the relevant user information or a default message if no data is returned
        # If you need to add first and last name, you'll need to add those to the both
        # returns (take in to account if nothing is found)
        return {
            "displayName": result.display_name,
            "mail": result.mail,
            "accountEnabled": result.account_enabled
        } if result else {
            "displayName": "No Data Returned",
            "mail": "No Data Returned",
            "accountEnabled": "No Data Returned"
        }
    except Exception as e:
        # Catch any exceptions and print the error message
        # Remember to also add additional fields for the returns as stated above
        error_message = str(e)
        print(
            Fore.RED + f"Error fetching user email for User ID {user_id}: {e}")

        # Handle specific error cases
        if "Request_ResourceNotFound" in error_message or "404" in error_message:
            return {
                "displayName": "User Not Found in Azure AD",
                "mail": "User Not Found in Azure AD",
                "accountEnabled": "User Not Found in Azure AD"
            }
        elif "403" in error_message or "Forbidden" in error_message:
            return {
                "displayName": "Access Denied",
                "mail": "Access Denied",
                "accountEnabled": "Access Denied"
            }
        elif "401" in error_message or "Unauthorized" in error_message:
            return {
                "displayName": "Authentication Failed",
                "mail": "Authentication Failed",
                "accountEnabled": "Authentication Failed"
            }
        else:
            return {
                "displayName": f"Error: {error_message[:50]}...",
                "mail": f"Error: {error_message[:50]}...",
                "accountEnabled": f"Error: {error_message[:50]}..."
            }


"""
=====================================================================================================================
End get_User_Email_Azure_python function
=====================================================================================================================
"""


"""
=====================================================================================================================
Load Requirements
=====================================================================================================================
"""

# Call the function to read the configuration file
config_data = read_config()


"""
=====================================================================================================================
Script start
=====================================================================================================================
"""

"""
=====================================================================================================================
Determine Environment to run against. 
=====================================================================================================================
"""

selected = False
while selected == False:
    # creates a prompt box to pick which system to use for the rest of the application
    msg = "What Kaltura system do you want to use?"
    title = "Choose System"
    choices = ["Test", "Prod", "Exit"]
    global envPicked
    envPicked = choicebox(msg, title, choices)

    # Take the picked environment and use the associated app ID and token to create a Kaltura session.
    match envPicked:
        case "Test":
            print(Fore.GREEN + "Test Selected")
            selected = True
        case "Prod":
            print(Fore.GREEN + "Prod Selected")
            selected = True
        case "Exit" | None:
            print(Fore.RED + "Exiting....")
            sys.exit()
            selected = True
        case _:
            print(Fore.RED + "Error with selection, try again")
            selected = False
"""
=====================================================================================================================
End Environment to run against. 
=====================================================================================================================
"""


# Main function to run the script
# Connect to Microsoft Graph API
try:
    graphClient = connectToMSGraph()
    print(Fore.GREEN + "Connected to Microsoft Graph API, now you can make requests!")
except Exception as e:
    print(Fore.RED + f"Error connecting to Microsoft Graph API: {e}")
    sys.exit(1)


print(Fore.GREEN + "Connection commpleted. ")
userEmail = get_User_Email_Azure_python("some user id is passed here")
print(Fore.GREEN + "Do something with the info retruned. Sleeping for a little beofre exiting")
time.sleep(sleepTimeSeconds)
print(Fore.RED + "Exiting....")
sys.exit()
