import openpyxl
import os
import requests
from shutil import copyfile

def main():
    
    # Error Checks
    path = os.getcwd()

    # Check for recent_reports folders
    if not os.path.exists(path + 'recent_reports'):
        pass

    # check for recent umcu and SOAS files
    if not os.path.isfile(path + '\\recent_reports\\umcu.xlsx'):
    #return print("File doesn't exits")
        pass
    if not os.path.isfile(path + '\\recent_reports\\soas.xlsx'):
    #return print("File doesn't exits")
        pass

    umcu = path + ''
    #copyfile(src, dst)

    # Download the bank statement from UMCU
    # Make session
    with requests.Session() as s:
        # Start at sign-in page
        signin_url = 'https://my.umcu.org/User/AccessSignin/Start'
        username = ''
        data = {'FormInstanceToken': '2DE044188F60428C92E16DA72DAAC94A', 'UsernameField': username, 'SubmitNext': 'Sign In'}
        response = s.post(signin_url, data=data, verify=True)
        print('Status code: ', response.status_code)
        print('Response URL: ' + response.url)
        # Step 2: Password page
        pass_url = 'https://my.umcu.org/User/AccessSignin/Password'
        password = ''
        data = {'FormInstanceToken': 'D1866717F71547E9B9E6817C360AD16A' ,'PasswordField': password, 'SubmitNext': 'Next'}
        response2 = s.post(pass_url, data=data, verify=True)
        print('Status code: ', response2.status_code)
        print('Response URL: ' + response2.url)
        # Step 3: Challenge Question
        challenge_url = 'https://my.umcu.org/User/AccessSignin/Challenge'
        answer = ''
        data = {'FormInstanceToken': '3DFD6FAF269340ED9D696FD23217E7F6', 'Answer': answer, 'Fizzbot': 'D8EBC11D78178119', 'Remember': 'False', 'SubmitNext': 'Next'}
        response3 = s.post(challenge_url, data=data, verify=True)
        print('Status code: ', response3.status_code)
        print('Response URL: ' + response3.url)
        # Step 4: eStatements
        estatement_url ='https://my.umcu.org/User/CustomSignonToEStatements/Start'
        response4 = s.get(estatement_url)
        print('Status code: ', response4.status_code)
        print('Response URL: ' + response4.url)
        # Step 5: Parse eStatement page for redirect link


    # Open the files



    # Extract the data


if __name__ == '__main__':
    main()