##############################-- Script Header --#################################
# Name:     Get-IncidentCount.py
# Version:  1.00.00
# Purpose:  To check if a provided list of email addresses are actually resolving
#           any tickets in ServiceNow
# Usage:    Use during license clean-up activity of ITIL Licenses from the instance
#               Pull out a list of users (email addresses) from sys_user_has_license table
#               paste the email addresses in exmail.xlsx
#               Make sure the patch to the email.xlsx is defined properly in the read_excel method
#               Execute the script.
# Output:   From the given list of email addresses, those users who have resolved 0 incidents
#           will be displayed on the console
# Args:     None
# History:  Version     Description
#           1.00.00     Initial release
# Author:   Tushar Singh
#################################################################################

########## References ###########
# 1. https://servicenow.github.io/PySNC/
# 2. https://servicenow.github.io/PySNC/user/getting_started.html#querying-a-record 
# 3. https://www.scaler.com/topics/pandas/pandas-read-excel/
# 4. https://www.geeksforgeeks.org/different-ways-to-iterate-over-rows-in-pandas-dataframe/
##################################

###### Imports ##########
from pysnc import ServiceNowClient
import pandas as pd
#########################

####### Begin Script Block ###########
# Create a client connection with ServiceNow instance. Ensure you have created a user on
# the desired instance and granted permissions
oClient = ServiceNowClient('<servicenow_instance_name>', ('<username>', '<password>'))

# Load the excel containing email addresses in a pandas data frame
df = pd.read_excel(io=r"<path to excel sheet>", sheet_name="Sheet1")

# Now we need the email addresses in the df one by one so that we can process one
# row at a time, fetching the "assigned_to" count for that user from the SN Instance
for row in df.index:
    #print(df['Email Address'][row])
    #element by element
    emailaddress = df['Email Address'][row]
    #object to the user table in servicenow
    gr_user = oClient.GlideRecord('sys_user')
    #add the filter to fetch only the specific user whose email address is at the
    #current row
    gr_user.add_query("email", emailaddress)
    #execute the query
    gr_user.query()
    #looping through all the user records that have come up from the query
    #usually only 1 record will come up since we are using email address as the
    #search filter.
    while gr_user.next():
        #print(gr_user.get_value('sys_id'))
        # In order to get this list of assigned incidents for this user
        # we would require the sys_id for the user record.
        sys_id = gr_user.get_value('sys_id')

        # For the email address (current row) we now have the sys_id. So now
        # we can create an object for the incident table and search through it
        # referencing the sys_id of the email address
        gr = oClient.GlideRecord('incident')
        gr.add_query("assigned_to",sys_id)
        gr.query()
        # Fetch the number of incidents ever assigned to this user that he / she resolved
        solvedcount = gr.get_row_count()

        # A condition to check only the ones who have not resolved anything
        if(solvedcount == 0):
            print(emailaddress, gr.get_row_count())

############# End of Script Block ###################################
 
