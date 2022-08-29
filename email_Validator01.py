import re
import smtplib
import dns.resolver
import pandas as pd
import csv

# Address used for SMTP MAIL FROM command  
fromAddress = 'corn@bt.com'

# Simple Regex for syntax checking
regex = '^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,})$'

# Email address to verify
fname = "mail20test.xlsx"
inputAddressList = []

df = pd.read_excel(fname)
for emailAdd in df['name']:
    inputAddressList.append(emailAdd)
    #print(emailAdd)

fw = open("validEmailsss.csv", 'w')
cw = csv.writer(fw)

for inputAddress in inputAddressList:
    code = 0
    row = []

    print(inputAddress)
    #inputAddress = input('Please enter the emailAddress to verify:')
    addressToVerify = str(inputAddress)

    # Syntax check
    match = re.match(regex, addressToVerify)
    if match == None:
        print('Bad Syntax')
        #raise ValueError('Bad Syntax')
    else:
        # Get domain for DNS lookup
        splitAddress = addressToVerify.split('@')
        domain = str(splitAddress[1])
        print('Domain:', domain)

        # MX record lookup
        try:
            records = dns.resolver.resolve(domain, 'MX')
            mxRecord = records[0].exchange
            mxRecord = str(mxRecord)


            # SMTP lib setup (use debug level for full output)
            server = smtplib.SMTP()
            server.set_debuglevel(0)

            # SMTP Conversation
            server.connect(mxRecord)
            server.helo(server.local_hostname) ### server.local_hostname(Get local server hostname)
            server.mail(fromAddress)
            code, message = server.rcpt(str(addressToVerify))
            server.quit()
        except :
            print(f'{domain}: Doesn\'t exist or No answer')

        #print(code)
        #print(message)

        # Assume SMTP response 250 is success
        if code == 250:
            print('Success')
            row.append(inputAddress) 
            cw.writerow(row)
        else:
            print('Bad')