from tkinter import *
import tkinter.scrolledtext as st
import win32com.client as wc
import xml.etree.ElementTree as ET
import pandas as pd
import mysql.connector as mc

# Get sql user and password
from config import mysql_host, mysql_u, mysql_pw
from config import vgc_host, vgc_u, vgc_pw

class VGCqbCommunicator():

    def __init__(self):
        # Create GUI Windows
        window = Tk()
        window.geometry("600x600")
        window.title("VGC QB Communicator")
        window.configure(background='#696773')
        window.iconbitmap('./images/Frostlogo_icon_32.ico')

        # Title label
        Label (window, text='VGC QB Communicator', bg='#696773', fg='#ececee', font=('Book Antiqua', 20, 'bold')) .grid(row=0, column=0, columnspan=4, padx = 30, pady = 10, sticky=W)

        # Add Frost Logo
        frostLogo = PhotoImage(file='./images/Frostlogo_icon.png')
        Label (window, image=frostLogo, bg='#696773') .grid(row=0, column=4, columnspan=2, padx = 30, pady = 10, sticky=W)

        # Status label
        Label (window, text='Status:', bg='#696773', fg='#ececee', font=('Book Antiqua', 13, 'bold')) .grid(row=1, column=0, padx = 10, pady = 5, sticky=W)

        # Status Scrolled text
        self.output = st.ScrolledText(window, width = 70, height = 8, wrap=WORD, background='#363946', fg='#ed9511', font=('Book Antiqua', 12, 'bold'))
        self.output.grid(row=3, column=0, columnspan=5, padx=(10,0))

        # Making the text read only
        # self.output.configure(state ='disabled')        

        self.statusText = 'Click a button to run a script.'
        self.output.insert(END, self.statusText)
        self.output.see(END)
        self.output.configure(state ='disabled')

        # Customer button
        self.customerBtn = Button(window, text='Get QB Customers', width=20, height=2, command=self.qbCustomers, bg='#1b1bd1', fg='#e3e3e6', font=('Book Antiqua', 13, 'bold')) 
        self.customerBtn.grid(row=5, column=0, columnspan=3, pady=20, padx=20, sticky=E) 
        self.customerBtn.bind("<Enter>", self.customerBtnEnter)
        self.customerBtn.bind("<Leave>", self.customerBtnClose)

        # Accounts button
        self.accountsBtn = Button(window, text='Get QB Accounts', width=20, height=2, command=self.qbAccounts, bg='#1b1bd1', fg='#e3e3e6', font=('Book Antiqua', 13, 'bold')) 
        self.accountsBtn.grid(row=5, column=3, columnspan=3, pady=20, padx=20, sticky=W) 
        self.accountsBtn.bind("<Enter>", self.accountsBtnEnter)
        self.accountsBtn.bind("<Leave>", self.accountsBtnClose)

        # VGC ==> QB button
        self.vgcToQbBtn = Button(window, text='VGC ==> QB', width=20, height=2, command=self.vgcToQb, bg='#1b1bd1', fg='#e3e3e6', font=('Book Antiqua', 13, 'bold')) 
        self.vgcToQbBtn.grid(row=6, column=0, columnspan=3, pady=20, padx=20, sticky=E) 
        self.vgcToQbBtn.bind("<Enter>", self.vgcToQbBtnEnter)
        self.vgcToQbBtn.bind("<Leave>", self.vgcToQbBtnClose)

        # QB ==> VGC button
        self.qbToVgcBtn = Button(window, text='QB ==> VGC', width=20, height=2, command=self.qbToVgc, bg='#1b1bd1', fg='#e3e3e6', font=('Book Antiqua', 13, 'bold')) 
        self.qbToVgcBtn.grid(row=6, column=3, columnspan=3, pady=20, padx=20, sticky=W) 
        self.qbToVgcBtn.bind("<Enter>", self.qbToVgcBtnEnter)
        self.qbToVgcBtn.bind("<Leave>", self.qbToVgcBtnClose)

        # Main loop
        window.mainloop()
    
    '''
        BUTTON HOVER FUNCTIONS
    '''
    def customerBtnEnter(self, sender):
        self.customerBtn['background']='#101078'

    def customerBtnClose(self, sender):
        self.customerBtn['background']='#1b1bd1'

    def accountsBtnEnter(self, sender):
        self.accountsBtn['background']='#101078'

    def accountsBtnClose(self, sender):
        self.accountsBtn['background']='#1b1bd1'

    def vgcToQbBtnEnter(self, sender):
        self.vgcToQbBtn['background']='#101078'

    def vgcToQbBtnClose(self, sender):
        self.vgcToQbBtn['background']='#1b1bd1'

    def qbToVgcBtnEnter(self, sender):
        self.qbToVgcBtn['background']='#101078'

    def qbToVgcBtnClose(self, sender):
        self.qbToVgcBtn['background']='#1b1bd1'

    # Status box functions
    def updateStatusText(self, text):
        self.output.configure(state ='normal')
        self.statusText = f'{text}'
        self.output.insert(END, self.statusText)
        self.output.update()
        self.output.see(END)
        self.output.configure(state ='disabled')

    def clearStatusText(self):
        self.output.configure(state = 'normal')
        self.output.delete(0.0,END)
        self.output.configure(state ='disabled')

    '''
        QB FUNCTIONS
    '''
    def qbToVgc(self):
        # Clear Output
        self.output.delete(0.0,END)

    def vgcToQb(self):
        # Clear Output
        self.output.delete(0.0,END)

        # Pull RTBP claim from VGC
        # connect to DB
        cnx = mc.connect(user=vgc_u, password=vgc_pw,
                        host=vgc_host,
                        database='visualgap_claims')
        cursor = cnx.cursor()

        # sql query for all accounts in DB
        sql_file = '''
                    SELECT c.claim_id, c.claim_nbr, c.carrier_id, cl.alt_name, cl.contact, cl.address1, 
                        cl.city, cl.state, cl.zip, cl.payment_method, cb.first, cb.last, 
                        IF(sq.gap_amt_paid > 0, 2,1) AS pymt_type_id, 
                        IF(sq.gap_amt_paid > 0, ROUND(cc.gap_payable - sq.gap_amt_paid,2), cc.gap_payable) AS gap_due
                    FROM claims c
                    INNER JOIN claim_lender cl
                        USING (claim_id)
                    INNER JOIN claim_borrower cb
                        USING (claim_id)
                    INNER JOIN claim_calculations cc
                        USING (claim_id)
                    INNER JOIN claim_status cs
                        ON (c.status_id = cs.status_id)
                    LEFT JOIN (SELECT cp.claim_id, SUM(cp.payment_amount) AS gap_amt_paid
                            FROM claim_payments cp
                            INNER JOIN (SELECT c.claim_id
                                        FROM claims c
                                        INNER JOIN claim_status cs
                                            ON (c.status_id = cs.status_id)
                                        WHERE cs.status_desc_id = 8) rtbp_sq
                                USING (claim_id)
                            WHERE payment_category_id = 1
                            GROUP BY cp.claim_id) sq
                        ON (c.claim_id = sq.claim_id)
                    WHERE cs.status_desc_id = 8;
                    '''

        # execute sql
        cursor.execute(sql_file)
        # save query results as DF
        df = pd.DataFrame(cursor.fetchall())
        # add column names
        df.columns=['claim_id', 'claim_nbr', 'carrier_id', 'lender_name', 'contact', 'address1', 'city', 'state', 'zip', 
                            'pymt_method', 'first', 'last', 'pymt_type_id', 'gap_due']

    def qbCustomers(self):
        # Clear Output
        self.clearStatusText()

        # Connect to Quickbooks
        statusText = f'Connecting to Quickbooks...'
        self.updateStatusText(statusText)

        sessionManager = wc.Dispatch("QBXMLRP2.RequestProcessor")    
        sessionManager.OpenConnection('', 'Test qbXML Request')
        ticket = sessionManager.BeginSession("", 2)

        # create qbxml query
        qbxmlQuery = '''
            <?qbxml version="14.0"?>
            <QBXML>
                <QBXMLMsgsRq onError="stopOnError">
                    <CustomerQueryRq requestID="1">
                    </CustomerQueryRq>
                </QBXMLMsgsRq>
            </QBXML>
            '''

        # Send query and receive response
        # self.output.configure(state ='normal')
        statusText = f'\r\nRunning query...'
        self.updateStatusText(statusText)

        responseString = sessionManager.ProcessRequest(ticket, qbxmlQuery)

        # Disconnect from Quickbooks
        # self.output.configure(state ='normal')
        statusText = f'\r\nDisconnecting from Quickbooks...'
        self.updateStatusText(statusText)

        sessionManager.EndSession(ticket)
        sessionManager.CloseConnection()

        # self.output FullName and ListID to a dataframe
        # create dataframe
        currentCustomerDF = pd.DataFrame(columns=['ListID','FullName'])
        
        # self.output.configure(state ='normal')
        statusText = f'\r\nParsing query response...'
        self.updateStatusText(statusText)

        QBXML = ET.fromstring(responseString)
        QBXMLMsgsRs = QBXML.find('QBXMLMsgsRs')
        customerResults = QBXMLMsgsRs.iter('CustomerRet')
        for customerResult in customerResults:
            customerListID = customerResult.find('ListID').text
            customerName = customerResult.find('FullName').text
            # add to dataframe
            currentCustomerDF = currentCustomerDF.append({'ListID': customerListID, 'FullName': customerName}, ignore_index=True)

        # self.output.configure(state ='normal')
        statusText = f'\r\nExporting results...'
        self.updateStatusText(statusText)

        # display df
        print(currentCustomerDF)
        #currentCustomerDF.to_csv('qbCustomersself.output.csv',header=True,index=False)
        # self.output.configure(state ='normal')
        statusText = f'\r\nComplete!'
        self.updateStatusText(statusText)

    def qbAccounts(self):
        # Clear Output
        self.clearStatusText()

        # Connect to Quickbooks
        statusText = f'Connecting to Quickbooks...'
        self.updateStatusText(statusText)

        sessionManager = wc.Dispatch("QBXMLRP2.RequestProcessor")    
        sessionManager.OpenConnection('', 'Test qbXML Request')
        ticket = sessionManager.BeginSession("", 2)

        # create qbxml query
        qbxmlQuery = '''
            <?qbxml version="14.0"?>
            <QBXML>
                <QBXMLMsgsRq onError="stopOnError">
                    <AccountQueryRq requestID="1">
                    </AccountQueryRq>
                </QBXMLMsgsRq>
            </QBXML>
            '''

        # Send query and receive response
        statusText = f'\r\nRunning query...'
        self.updateStatusText(statusText)

        responseString = sessionManager.ProcessRequest(ticket, qbxmlQuery)

        # Disconnect from Quickbooks
        statusText = f'\r\nDisconnecting from Quickbooks...'
        self.updateStatusText(statusText)

        sessionManager.EndSession(ticket)
        sessionManager.CloseConnection()

        # create dataframe to store QB accounts
        currentAccountDF = pd.DataFrame(columns=['qb_listid','qb_fullname'])

        statusText = f'\r\nParsing query response...'
        self.updateStatusText(statusText)

        # self.output FullName and ListID to a dataframe
        QBXML = ET.fromstring(responseString)
        QBXMLMsgsRs = QBXML.find('QBXMLMsgsRs')
        accountResults = QBXMLMsgsRs.iter("AccountRet")
        for accountResult in accountResults:
            accountListID = accountResult.find('ListID').text
            accountName = accountResult.find('FullName').text
            # add to dataframe
            currentAccountDF = currentAccountDF.append({'qb_listid': accountListID, 'qb_fullname': accountName}, ignore_index=True)

        # connect to DB
        statusText = f'\r\nConnecting to MySQL database...'
        self.updateStatusText(statusText)

        cnx = mc.connect(user=mysql_u, password=mysql_pw,
                        host=mysql_host,
                        database='claim_qb_payments')
        cursor = cnx.cursor()

        # sql query for all accounts in DB
        statusText = f'\r\nRunning SQL query...'
        self.updateStatusText(statusText)

        sql_file = '''
            SELECT qb_listid, qb_fullname
            FROM qb_accounts;
            '''

        # execute sql
        cursor.execute(sql_file)
        # save query results as DF
        df = pd.DataFrame(cursor.fetchall())

        # add column names to DF
        col_names = ['qb_listid', 'qb_fullname']
        df.columns = col_names

        # check for new accounts since last update
        statusText = f'\r\nChecking for new accounts...'
        self.updateStatusText(statusText)

        if (len(df) != len(currentAccountDF)):

            # use join to find new accounts
            checkNewAccountsDF = pd.merge(currentAccountDF, df, how='left', indicator=True).copy()
            newAccountsDF = checkNewAccountsDF[checkNewAccountsDF['_merge'].eq('left_only')].drop(['_merge'], axis=1).copy()

            # add carrier_id and account_type columns then reorder to match db
            newAccountsDF['carrier_id'] = 0
            newAccountsDF['account_type'] = ''
            newAccountsDF = newAccountsDF.reindex(columns=['carrier_id', 'qb_listid', 'qb_fullname', 'account_type']).copy()

            statusText = f'\r\nAdd new accounts...'
            self.updateStatusText(statusText)

            for i,row in newAccountsDF.iterrows():
                # display message to add carrier ID (from VG) and update df
                newAccountsDF.at[i,'carrier_id'] = input(f"Enter the VG Carrier ID for ''{row['qb_fullname']}'' (ANICO = 9 & Securian = 8):")

                # display message to add type (Checking, Expense) and update df
                newAccountsDF.at[i,'account_type'] = input(f"Enter the account type for ''{row['qb_fullname']}'' (Checking or Expense):")
        
            # add new account to sql db
            # collect headers
            cols = ", ".join([str(i) for i in newAccountsDF.columns.tolist()])

            # create sql query
            for x,rows in newAccountsDF.iterrows():
                sql = f"INSERT INTO qb_accounts ({cols}) VALUES ({rows['carrier_id']}, '{rows['qb_listid']}', '{rows['qb_fullname']}', '{rows['account_type']}');" 

                # execute and commit sql
                cursor.execute(sql)
                cnx.commit()

        else:
            statusText = f'\r\nNo new accounts found...'
            self.updateStatusText(statusText)
    
        # close mysql connection
        statusText = f'\r\nClosing MySQL connection...'
        self.updateStatusText(statusText)

        cursor.close()
        cnx.close()

        statusText = f'\r\nComplete!'
        self.updateStatusText(statusText)

def main():
    VGCqbCommunicator()

if __name__ == '__main__':
    main()
