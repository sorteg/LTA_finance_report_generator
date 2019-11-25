import openpyxl
import os
from shutil import copyfile

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
copyfile(src, dst)

# Download the bank statement from UMCU
# Make session?

# Start at sign-in page
signin_url = 'https://my.umcu.org/User/AccessSignin/Start'
username = 'bolta1975'
# send GET request

# Step 2: Password page
pass_url = 'https://my.umcu.org/User/AccessSignin/Password'
password = 'palma1975'
# send GET request

# Step 3: Challenge Question
challenge_url = 'https://my.umcu.org/User/AccessSignin/Challenge'
answer = '1975'
# send GET request

# Step 4: eStatements
estatement_url = 'https://my.umcu.org/User/CustomSignonToEStatements/Start'
# handle redirect?

# Open the files



# Extract the data