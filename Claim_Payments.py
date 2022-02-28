from tkinter import *
import tkinter.scrolledtext as st
import win32com.client as wc
import xml.etree.ElementTree as ET
import pandas as pd
import mysql.connector as mc
import datetime as dt
from datetime import datetime
import pyodbc
from pdfrw import PdfReader, PdfWriter
import pdfrw 
import os
from pdf2image import convert_from_path
import img2pdf
from PIL import Image
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import sys
import shutil
from bs4 import BeautifulSoup as Soup
import numpy as np
import pdfkit
import pysftp

# Get credentials
from configTest import mysql_host, mysql_u, mysql_pw, vgc_host, vgc_u, vgc_pw, svr, db, sql_u, sql_pw, smtp_host, e_user, e_pw, port, sftp_h, sftp_u, sftp_p
# from config import mysql_host, mysql_u, mysql_pw, vgc_host, vgc_u, vgc_pw

class VGCqbCommunicator():

    def __init__(self):
        # Create GUI Windows
        window = Tk()
        window.geometry("600x600")
        window.title("Claim Payments")
        window.configure(background='#696773')
        window.iconbitmap('./images/Frostlogo_icon_32.ico')

        # Title label
        Label (window, text='Claim Payments', bg='#696773', fg='#ececee', font=('Book Antiqua', 30, 'bold')) .grid(row=0, column=0, columnspan=4, padx = 30, pady = 10, sticky=W)

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

        # Define directories
        self.attachment_dir = 'S:/claims/letters/attachment/'
        self.file_staging_dir = './letters/staging/'
        # get current date
        self.now = datetime.now()

        # Main loop
        window.mainloop()
    
    #===========================
    #    BUTTON HOVER FUNCTIONS
    #===========================
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
    
    #===========================
    #   OPERATIONAL FUNCTIONS
    #===========================

    # MySQL
    def mysql_q (self, u, p, h, db, sql, cols, commit):
        # cols (0 = no, 1 = yes)
        # commit (select = 0, insert/update = 1)

        # connect to claim_qb_payments db
        cnx = mc.connect(user=u, password=p,
                        host=h,
                        database=db)
        cursor = cnx.cursor()

        # commit?
        if commit == 1:
            cursor.execute(sql)
            cnx.commit()
            sql_result = 0     
        else:
            cursor.execute(sql)
            sql_result = cursor.fetchall()

        # columns ?
        if cols == 1:
            columns=list([x[0] for x in cursor.description])
            # close connection
            cursor.close()
            cnx.close()
            # return query result [0] and columns [1]
            return sql_result
        else:
            # close connection
            cursor.close()
            cnx.close()
            # return query result
            return sql_result

    # PDF Concatenation Function
    def ConCat_pdf (self, file_list, outfn):
        letter_path = './letters/staging/'
        writer = PdfWriter()
        for inputfn in file_list:
            writer.addpages(PdfReader(letter_path + inputfn).pages)

        outfile = outfn + '.pdf'
        fnNum = 0

        while (os.path.isfile(outfile) == True):
            fnNum += 1
            outfile = outfn + '-' + str(fnNum) + '.pdf'

        writer.write(outfile)
        return outfile

    # Delete File Function
    def delete_file(self, del_file_path):
        if os.path.exists(del_file_path):
            os.remove(del_file_path)
        else: print (f"{del_file_path} does not exist")

    # Create PDF Function
    def fill_pdf(self, input_pdf_path, output_pdf_path, data_dict):
        ANNOT_KEY = '/Annots'
        ANNOT_FIELD_KEY = '/T'
        ANNOT_VAL_KEY = '/V'
        ANNOT_RECT_KEY = '/Rect'
        SUBTYPE_KEY = '/Subtype'
        WIDGET_SUBTYPE_KEY = '/Widget'

        template_pdf = pdfrw.PdfReader(input_pdf_path)
        
        for page in template_pdf.pages:
            annotations = page[ANNOT_KEY]
            for annotation in annotations:
                if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                    if annotation[ANNOT_FIELD_KEY]:
                        key = annotation[ANNOT_FIELD_KEY][1:-1]
                        if key in data_dict.keys():
                            if type(data_dict[key]) == bool:
                                if data_dict[key] == True:
                                    annotation.update(pdfrw.PdfDict(
                                        AS=pdfrw.PdfName('Yes')))
                            else:
                                annotation.update(
                                    pdfrw.PdfDict(V='{}'.format(data_dict[key]))
                                )
                                annotation.update(pdfrw.PdfDict(AP=''))
        template_pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))

        pdfrw.PdfWriter().write(output_pdf_path, template_pdf)

    # Flatten PDF Function
    def flatten_pdf(self, flat_output, img_file):
        # Fillable PDF to Image
        images = convert_from_path(flat_output, dpi=300, size=(2550,3300))
        for i in range(len(images)):
    
        # Save pages as images in the pdf
            images[i].save(img_file + '.png', 'PNG')
        
        # Delete Fillable PDF
        self.delete_file(img_file + '.pdf')
        
        # opening image
        image_file = Image.open(img_file + '.png')
        
        # Image to Flat PDF
        # define paper size
        letter = (img2pdf.in_to_pt(8.5), img2pdf.in_to_pt(11))
        layout = img2pdf.get_layout_fun(letter)
        # converting into chunks using img2pdf
        pdf_bytes = img2pdf.convert(image_file.filename, layout_fun=layout)
        
        # opening or creating pdf file
        flat_pdf = f"{img_file}.pdf"
        file = open(flat_pdf, "wb")
        
        # writing pdf files with chunks
        file.write(pdf_bytes)

        # closing image file
        image_file.close()
        
        # Delete Fillable PDF
        self.delete_file(img_file + '.png')

        # closing pdf file
        file.close()

    # GAP Letter Function
    def gap_letter(self, template_df, pdf_template, position):
        gap_path = './letters/staging/'

        template_df['payment_amount'] = template_df['payment_amount'].map('${:,.2f}'.format)
        template_df['loss_date'] = pd.to_datetime(template_df['loss_date']).dt.strftime('%B %d, %Y')
        template_df['StateDesc'] = template_df['StateDesc'].astype(str).replace({'None':''})
        template_df['StateCode'] = template_df['StateCode'].astype(str).replace({'None':''})
        template_df['f_lang'] = template_df['f_lang'].astype(str).replace({'None':''})
        letter_date = f"{datetime.now():%B %d, %Y}"   

        for index, row in template_df.iterrows():

            # empty dict
            data_dict = {}
            # store field data in dictionary
            data_dict = {
                'Date': letter_date,
                'Lender': template_df.loc[index]['alt_name'],
                'Contact': template_df.loc[index]['contact'],
                'Address': template_df.loc[index]['address1'],
                'City_St_Zip': f"{template_df.loc[index]['city']}, {template_df.loc[index]['state']} {template_df.loc[index]['zip']}",
                'Lender2': template_df.loc[index]['alt_name'],
                'Borrower': f"{template_df.loc[index]['first']} {template_df.loc[index]['last']}",
                'Claim_Nbr': template_df.loc[index]['claim_nbr'],
                'Acct_Nbr': template_df.loc[index]['acct_number'],
                'DOL': template_df.loc[index]['loss_date'],
                'GAP_Amt': template_df.loc[index]['payment_amount'],
                'State': template_df.loc[index]['StateDesc'],
                'St_Code': template_df.loc[index]['StateCode'],
                'Fraud': template_df.loc[index]['f_lang'],
            }

            # store paths as variables
            output_file = f"{template_df.loc[index]['claim_nbr']}-{position}.pdf"
            output_path_fn = f"{gap_path}{output_file}"

            self.fill_pdf(pdf_template, output_path_fn, data_dict)

            # Set File Paths
            flat_output = f"{os.path.dirname(os.path.abspath(output_file))}\{gap_path}\{output_file}"
            img_file = f"{os.path.dirname(os.path.abspath(output_file))}\{gap_path}\{template_df.loc[index]['claim_nbr']}-{position}"

            # Flatten pdf using flatten_pdf function
            self.flatten_pdf(flat_output, img_file) 

    # GAP Calculation Function
    def calculations(self, template_df, pdf_template, position):
        gap_path = './letters/staging/'

        template_df['loss_date'] = pd.to_datetime(template_df['loss_date']).dt.strftime('%B %d, %Y')
        template_df['last_payment'] = pd.to_datetime(template_df['last_payment']).dt.strftime('%B %d, %Y')  
        template_df['incp_date'] = pd.to_datetime(template_df['incp_date']).dt.strftime('%B %d, %Y')    
        template_df['payoff'] = template_df['payoff'].map('${:,.2f}'.format)
        template_df['past_due'] = template_df['past_due'].map('${:,.2f}'.format)
        template_df['late_fees'] = template_df['late_fees'].map('${:,.2f}'.format)
        template_df['skip_pymts'] = template_df['skip_pymts'].map('${:,.2f}'.format)
        template_df['skip_fees'] = template_df['skip_fees'].map('${:,.2f}'.format)
        template_df['primary_pymt'] = template_df['primary_pymt'].map('${:,.2f}'.format)
        template_df['excess_deductible'] = template_df['excess_deductible'].map('${:,.2f}'.format)
        template_df['scr'] = template_df['scr'].map('${:,.2f}'.format)
        template_df['clr'] = template_df['clr'].map('${:,.2f}'.format)
        template_df['cdr'] = template_df['cdr'].map('${:,.2f}'.format)
        template_df['oref'] = template_df['oref'].map('${:,.2f}'.format)
        template_df['salvage'] = template_df['salvage'].map('${:,.2f}'.format)
        template_df['prior_dmg'] = template_df['prior_dmg'].map('${:,.2f}'.format)
        template_df['over_ltv'] = template_df['over_ltv'].map('${:,.2f}'.format)
        template_df['other1_amt'] = template_df['other1_amt'].map('${:,.2f}'.format)
        template_df['other2_amt'] = template_df['other2_amt'].map('${:,.2f}'.format)
        template_df['gap_payable'] = template_df['gap_payable'].map('${:,.2f}'.format)
        template_df['balance_last_pay'] = template_df['balance_last_pay'].map('${:,.2f}'.format)
        template_df['per_day'] = template_df['per_day'].map('${:,.2f}'.format)
        template_df['deductible'] = template_df['deductible'].map('${:,.2f}'.format)
        template_df['subtotal'] = template_df['subtotal'].map('${:,.2f}'.format)
        template_df['nbr_of_days'] = template_df['nbr_of_days'].map('{:,.0f}'.format)
        template_df['interest_rate'] = template_df['interest_rate'].map('{:,.2f}%'.format)
        template_df['ltv'] = template_df['ltv'].map('{:,.2f}%'.format)
        template_df['ltv_limit'] = template_df['ltv_limit'].map('{:,.2f}%'.format)
        template_df['percent_uncovered'] = template_df['percent_uncovered'].map('{:,.2f}%'.format)
        template_df['covered_fin_amount'] = template_df['covered_fin_amount'].map('${:,.2f}'.format)
        template_df['Amt_Fin'] = template_df['Amt_Fin'].map('${:,.2f}'.format)
        template_df['nada_value'] = template_df['nada_value'].map('${:,.2f}'.format)

        for index, row in template_df.iterrows():

            # empty dict
            data_dict = {}
            # store field data in dictionary
            data_dict = {
                'Claim_Number': template_df.loc[index]['claim_nbr'],
                'Status': 'Paid',
                'Borrower': f"{template_df.loc[index]['first']} {template_df.loc[index]['last']}",
                'Vehicle': template_df.loc[index]['vehicle'],
                'Date_Of_Loss': template_df.loc[index]['loss_date'],
                'Type_Of_Loss': template_df.loc[index]['loss_type'],
                'Lender': template_df.loc[index]['alt_name'],         
                'Lender_Contact': template_df.loc[index]['contact'],
                'Insurance_Carrier': template_df.loc[index]['carrier'],
                'Inception_Date': template_df.loc[index]['incp_date'],
                'Deductible': template_df.loc[index]['deductible'],
                'Payoff': template_df.loc[index]['payoff'],
                'Past_Due': template_df.loc[index]['past_due'],
                'Late_Fees': template_df.loc[index]['late_fees'],
                'Skips': template_df.loc[index]['skip_pymts'],
                'Skip_Fees': template_df.loc[index]['skip_fees'],
                'Primary': template_df.loc[index]['primary_pymt'],
                'Deductible_Excess': template_df.loc[index]['excess_deductible'],
                'SCR': template_df.loc[index]['scr'],
                'CL_Refund': template_df.loc[index]['clr'],
                'CD_Refund': template_df.loc[index]['cdr'],
                'O_Refund': template_df.loc[index]['oref'],
                'Salvage': template_df.loc[index]['salvage'],
                'Prior_Damage': template_df.loc[index]['prior_dmg'],
                'Over_LTV': template_df.loc[index]['over_ltv'],
                'Other1_Description': template_df.loc[index]['other1_description'],
                'Other2_Description': template_df.loc[index]['other2_description'],
                'Other1': template_df.loc[index]['other1_amt'],
                'Other2': template_df.loc[index]['other2_amt'],
                'Deduction_Subtotal': template_df.loc[index]['subtotal'],
                'GAP_Amt': template_df.loc[index]['gap_payable'], 
                'Last_pymt_date': template_df.loc[index]['last_payment'], 
                'DOL': template_df.loc[index]['loss_date'],
                'Number_of_days': template_df.loc[index]['nbr_of_days'], 
                'Loan_Payoff_As_of_DOL': template_df.loc[index]['payoff'],
                'Bal_as_of_last_pymt': template_df.loc[index]['balance_last_pay'],
                'Interest_Rate': template_df.loc[index]['interest_rate'],
                'Interest_Per_Day': template_df.loc[index]['per_day'],
                'Amt_financed': template_df.loc[index]['Amt_Fin'],
                'ACV': template_df.loc[index]['nada_value'], 
                'LTV': template_df.loc[index]['ltv'],
                'Max_Amt_Financed': template_df.loc[index]['covered_fin_amount'],
                'LTV_limit': template_df.loc[index]['ltv_limit'],
                'Percentage_Not_Covered': template_df.loc[index]['percent_uncovered']
            }

            # store paths as variables
            output_file = f"{template_df.loc[index]['claim_nbr']}-{position}.pdf"
            output_path_fn = f"{gap_path}{output_file}"

            self.fill_pdf(pdf_template, output_path_fn, data_dict)

            # Set File Paths
            flat_output = f"{os.path.dirname(os.path.abspath(output_file))}\{gap_path}\{output_file}"
            img_file = f"{os.path.dirname(os.path.abspath(output_file))}\{gap_path}\{template_df.loc[index]['claim_nbr']}-{position}"

            # Flatten pdf using flatten_pdf function
            self.flatten_pdf(flat_output, img_file)

    # Delete Files From Directory Function
    def clear_dir (self, path, ext):
        for x in os.listdir(path):
            if x.endswith(ext):
                os.remove(os.path.join(path, x))

    # Create File List Function
    def fileList (self, path, ext):
        file_list = []
        for x in os.listdir(path):
            if x.endswith(ext):
                file_list.append(x)

        # sort list to collate pages
        file_list.sort()
        
        return file_list

    # TotalRestart Calculation Function
    def tr_calculations(self, template_df, pdf_template, position):
        gap_path = './letters/staging/'

        template_df['loss_date'] = pd.to_datetime(template_df['loss_date']).dt.strftime('%B %d, %Y')
        template_df['incp_date'] = pd.to_datetime(template_df['incp_date']).dt.strftime('%B %d, %Y')
        template_df['primary_pymt'] = template_df['primary_pymt'].map('${:,.2f}'.format)
        template_df['excess_deductible'] = template_df['excess_deductible'].map('${:,.2f}'.format)
        template_df['scr'] = template_df['scr'].map('${:,.2f}'.format)
        template_df['clr'] = template_df['clr'].map('${:,.2f}'.format)
        template_df['cdr'] = template_df['cdr'].map('${:,.2f}'.format)
        template_df['oref'] = template_df['oref'].map('${:,.2f}'.format)
        template_df['salvage'] = template_df['salvage'].map('${:,.2f}'.format)
        template_df['prior_dmg'] = template_df['prior_dmg'].map('${:,.2f}'.format)
        template_df['other1_amt'] = template_df['other1_amt'].map('${:,.2f}'.format)
        template_df['other2_amt'] = template_df['other2_amt'].map('${:,.2f}'.format)
        template_df['other3_amt'] = template_df['other3_amt'].map('${:,.2f}'.format)    
        template_df['gap_payable'] = template_df['gap_payable'].map('${:,.2f}'.format)
        template_df['subtotal'] = template_df['subtotal'].map('${:,.2f}'.format)
        template_df['nada_value'] = template_df['nada_value'].map('${:,.2f}'.format)
        template_df['max_benefit'] = template_df['max_benefit'].map('${:,.2f}'.format)    
        template_df['totalrestart_payable'] = template_df['totalrestart_payable'].map('${:,.2f}'.format)

        for index, row in template_df.iterrows():

            # empty dict
            data_dict = {}
            # store field data in dictionary
            data_dict = {
                'Claim_Number': template_df.loc[index]['claim_nbr'],
                'Status': 'Paid',
                'Borrower': f"{template_df.loc[index]['first']} {template_df.loc[index]['last']}",
                'Vehicle': template_df.loc[index]['vehicle'],
                'Date_Of_Loss': template_df.loc[index]['loss_date'],
                'Type_Of_Loss': template_df.loc[index]['loss_type'],
                'Lender': template_df.loc[index]['alt_name'],         
                'Lender_Contact': template_df.loc[index]['contact'],
                'Inception_Date': template_df.loc[index]['incp_date'],
                'Max_Potential_Benefit': template_df.loc[index]['max_benefit'],
                'Membership_Term': template_df.loc[index]['term'],
                'Primary': template_df.loc[index]['primary_pymt'],
                'Deductible_Excess': template_df.loc[index]['excess_deductible'],
                'SCR': template_df.loc[index]['scr'],
                'CL_Refund': template_df.loc[index]['clr'],
                'CD_Refund': template_df.loc[index]['cdr'],
                'O_Refund': template_df.loc[index]['oref'],
                'Salvage': template_df.loc[index]['salvage'],
                'Prior_Damage': template_df.loc[index]['prior_dmg'],
                'Other1_Description': template_df.loc[index]['other1_description'],
                'Other2_Description': template_df.loc[index]['other2_description'],
                'Other3_Description': template_df.loc[index]['other3_description'],
                'Other1': template_df.loc[index]['other1_amt'],
                'Other2': template_df.loc[index]['other2_amt'],
                'Other3': template_df.loc[index]['other3_amt'],
                'TR_Deduction_Subtotal': template_df.loc[index]['subtotal'],
                'GAP_Amt': template_df.loc[index]['gap_payable'], 
                'ACV': template_df.loc[index]['nada_value'], 
                'TR_Amt': template_df.loc[index]['totalrestart_payable']
            }

            # store paths as variables
            output_file = f"{template_df.loc[index]['claim_nbr']}-{position}.pdf"
            output_path_fn = f"{gap_path}{output_file}"

            self.fill_pdf(pdf_template, output_path_fn, data_dict)

            # Set File Paths
            flat_output = f"{os.path.dirname(os.path.abspath(output_file))}\{gap_path}\{output_file}"
            img_file = f"{os.path.dirname(os.path.abspath(output_file))}\{gap_path}\{template_df.loc[index]['claim_nbr']}-{position}"

            # Flatten pdf using flatten_pdf function
            self.flatten_pdf(flat_output, img_file)

    # Email Function
    def send_email(self, toEmail, subject, msg_html, attachPath='', *args):
        fromEmail = 'claims@visualgap.com'

        # address message
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = fromEmail
        msg['To'] = ','.join(toEmail)

        # create body
        body_html = MIMEText(msg_html, 'html')
        msg.attach(body_html) 

        for arg in args:
            fName = attachPath + arg
            filename = os.path.abspath(fName)

            with open(filename, 'rb') as fn:
                attachment = MIMEApplication(fn.read())
                attachment.add_header('Content-Disposition', 'attachment', filename=arg)
                msg.attach(attachment)

        context = ssl.create_default_context()
        try:
            server = smtplib.SMTP(smtp_host, port)
            # check connection
            server.ehlo()  
            # Secure the connection
            server.starttls(context=context)  
            # check connection
            server.ehlo()
            server.login(e_user, e_pw)
            # Send email
            server.sendmail(fromEmail, toEmail, msg.as_string())

        except Exception as e:
            # Print any error messages
            print(e)
        finally:
            server.quit()

    # Update toVGC Function
    def update_tovgc_1(self, df):
        # Update toVGC to 1

        if len(df) > 0:
            for index, row in df.iterrows():
                err_sql = '''
                        UPDATE ready_to_be_paid
                        SET toVGC = 1
                        WHERE rtbp_id = {rtbp_id};
                        '''.format(rtbp_id=row['rtbp_id'])

                # connect to claim_qb_payments db
                self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', err_sql, 0, 1)

    #===========================
    #    QB FUNCTIONS
    #===========================
    def qbToVgc(self):
        # Clear Output
        self.clearStatusText()

        # Define file paths
        attach_name = f'Claims_Paid_{self.now.strftime("%Y-%m-%d")}'
        attach_file = f'{self.attachment_dir}{attach_name}'
        html_file = f'{self.file_staging_dir}final_rpt.html'
        err_html_file = f'{self.file_staging_dir}error.html'
        csv_file = f'{self.file_staging_dir}final_rpt.csv'
        err_csv_file = f'{self.file_staging_dir}error.csv'
        template_file = './html/pymt_summary_template.html'
        html_file2 = './letters/staging/Claims_Paid.html'

        # sql query to collect qb_txnid's
        sql = '''
            SELECT rtbp_id, check_nbr, qb_txnid
            FROM ready_to_be_paid
            WHERE toVGC = 1
                AND qb_txnid <> '0';
            '''

        # save query results as DF
        statusText = f'Collecting recent claim payment records...'
        self.updateStatusText(statusText)

        try:
            qb_df = pd.DataFrame(self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql, 0, 0))

            # add column names
            qb_df_cols = ['rtbp_id','check_nbr', 'qb_txnid']
            qb_df.columns = qb_df_cols
        except Exception as e:
            print('''
            There are not any claim payments update.
            ERROR: {}'''.format(e))

            statusText = f'\r\nThere are not any claim payments to update.\r\nProcess Cancelled.'
            self.updateStatusText(statusText)

            return        

        # Connect to Quickbooks
        statusText = f'\r\nConnecting to Quickbooks...'
        self.updateStatusText(statusText)

        try:
            sessionManager = wc.Dispatch("QBXMLRP2.RequestProcessor")    
            sessionManager.OpenConnection('', 'Claim Payments')
            ticket = sessionManager.BeginSession("", 2)
        except Exception as e:
            print('''
            Make sure QuickBooks is running and you are logged into the Company File.
            ERROR: {}'''.format(e))

            statusText = f'\r\nThere was an ERROR connecting to QuickBooks.\r\n***Make sure QuickBooks desktop is open and you are logged into the correct Company file.***\r\nProcess Cancelled.'
            self.updateStatusText(statusText)

            return

        # create qbxml to query qb for check numbers
        statusText = f'\r\nCollecting data from Quickbooks...'
        self.updateStatusText(statusText)        
        
        try:
            for index, row in qb_df.iterrows():
                qbxmlQuery = '''
                            <?qbxml version="14.0"?>
                            <QBXML>
                                <QBXMLMsgsRq onError="stopOnError">
                                    <CheckQueryRq>
                                        <TxnID>{txnId}</TxnID> 
                                    </CheckQueryRq>
                                </QBXMLMsgsRq>
                            </QBXML>
                            '''.format(txnId=row['qb_txnid'])

                # Send query and receive response
                responseString = sessionManager.ProcessRequest(ticket, qbxmlQuery)

                # output Check Number (RefNumber)
                QBXML = ET.fromstring(responseString)
                QBXMLMsgsRs = QBXML.find('QBXMLMsgsRs')
                checkResults = QBXMLMsgsRs.iter("CheckRet")
                chkNbr = '0'
                for checkResult in checkResults:
                    chkNbr = checkResult.find('RefNumber').text

                # Add Check Number to ready_to_be_paid table
                qb_sql_file = '''UPDATE ready_to_be_paid
                            SET check_nbr = '{ChkNbr}'
                            WHERE rtbp_id = {rowID};'''.format(ChkNbr=chkNbr, rowID=row['rtbp_id'])
                
                # execute and commit sql
                self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', qb_sql_file, 0, 1)
        except Exception as e:
            print('''
            Make sure to print checks and process ACH in Quickbooks prior to starting this process.
            ERROR: {}'''.format(e))

            statusText = f'\r\nThere was an ERROR connecting to Quickbooks.\r\n***Make sure you ahve Quickbooks desktop open***\r\nProcess Cancelled.'
            self.updateStatusText(statusText)

            return

        # Disconnect from Quickbooks
        statusText = f'\r\nDisconnecting from Quickbooks...'
        self.updateStatusText(statusText)

        sessionManager.EndSession(ticket)
        sessionManager.CloseConnection()

        # sql query for GAP claims that are RTBP
        sql = '''
            SELECT r.rtbp_id, r.claim_id, r.claim_nbr, r.carrier_id, r.lender_name, r.pymt_method, r.first, r.last, r.pymt_type_id, r.amount, r.payment_category_id, r.check_nbr, r.qb_txnid, r.pymt_date, r.toVGC, r.err_msg
            FROM ready_to_be_paid r
            INNER JOIN (SELECT batch_id
                        FROM ready_to_be_paid
                        WHERE toVGC = 1
                        GROUP BY batch_id) sq
            USING(batch_id);
            '''

        # save query results as DF
        statusText = f'\r\nCollecting claim payment data...'
        self.updateStatusText(statusText)

        df = pd.DataFrame(self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql, 0, 0))

        # add column names
        df_cols = ['rtbp_id', 'claim_id', 'claim_nbr', 'carrier_id', 'lender_name', 'pymt_method', 'first', 'last', 'pymt_type_id', 'amount', 'payment_category_id', 'check_nbr', 'qb_txnid', 'pymt_date', 'toVGC', 'err_msg']
        df.columns = df_cols

        # Add carrier name
        sql = '''
            SELECT carrier_id, description
            FROM carriers;
            '''
        # save query results as DF
        carrier_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql, 0, 0))

        col_names = ['carrier_id', 'carrier']
        carrier_df.columns = col_names

        # Merge QB_ListID into df
        df = df.merge(carrier_df, left_on='carrier_id', right_on='carrier_id').copy()

        # format df
        statusText = f'\r\nFormatting claim payment data...'
        self.updateStatusText(statusText)        
        # list of conditions
        type_conds = [((df['payment_category_id'] == 1) & (df['pymt_type_id'] == 1)),
                    ((df['payment_category_id'] == 1) & (df['pymt_type_id'] == 2)),
                    ((df['payment_category_id'] == 2) & (df['pymt_type_id'] == 1)),
                    ((df['payment_category_id'] == 2) & (df['pymt_type_id'] == 2)),
                    ((df['payment_category_id'] == 3) & (df['pymt_type_id'] == 1)),
                    ((df['payment_category_id'] == 3) & (df['pymt_type_id'] == 2))]
        # list of name types
        type_name = ['GAP', 'GAP Supp', 'GAP Plus', 'GAP Plus Supp', 'TotalRestart', 'TotalRestart Supp']

        # add column and assigned values
        df['Claim_Type'] = np.select(type_conds, type_name)

        # Create df to update VGC
        # vgc_update_df = df[['claim_id', 'pymt_method', 'pymt_type_id', 'amount', 'payment_category_id', 'check_nbr', 'qb_txnid', 'pymt_date']].copy()

        # format amount
        df['amount'] = df['amount'].map('${:,.2f}'.format)

        # Create df for Claim Payment Summary
        final_rpt_df = df[['claim_nbr', 'carrier', 'lender_name', 'first', 'last', 'amount', 'pymt_method', 'check_nbr', 'pymt_date', 'Claim_Type']].loc[df['toVGC'] == 1].copy()

        # rename columns
        final_rpt_df.rename(columns = {'claim_nbr':'Claim Nbr', 'carrier':'Carrier', 'lender_name':'Lender', 'first':'First Name', 'last':'Last Name', 'amount':'Amount', 'pymt_method':'Method',
                                    'check_nbr':'Check Nbr', 'pymt_date': 'Date', 'Claim_Type':'Claim Type'}, inplace=True)

        # sort df
        final_rpt_df.sort_values(by = ['Carrier', 'Last Name', 'First Name'], inplace=True)

        # create error_df
        statusText = f'\r\nChecking claim payment records for errors...'
        self.updateStatusText(statusText)

        error_df = df[['claim_nbr', 'lender_name', 'first', 'last', 'err_msg']].loc[df['toVGC'] == 2].copy()

        # Format df
        error_df.rename(columns = {'claim_nbr':'Claim Nbr', 'lender_name':'Lender', 'first':'First Name', 'last':'Last Name', 'err_msg':'Error Message'}, inplace=True)

        # export is df as csv
        statusText = f'\r\nCreating Claim Payment Summary report...'
        self.updateStatusText(statusText)

        final_rpt_df.to_csv(csv_file, index=False)
        error_df.to_csv(err_csv_file, index=False)
        # read in csv
        csvFile = pd.read_csv(csv_file)
        errCsvFile = pd.read_csv(err_csv_file)
        # convert csv to html
        csvFile.to_html(html_file, index=False)
        errCsvFile.to_html(err_html_file, index=False)

        # Get html df
        soup = Soup(open(html_file), "html.parser")
        err_soup = Soup(open(err_html_file), "html.parser")
        table = str(soup.select_one("table", {"class":"dataframe"}))
        err_table = str(err_soup.select_one("table", {"class":"dataframe"}))

        # Get template
        soup2 = Soup(open(template_file), "html.parser")
        # Find and insert payment table
        df_div = soup2.find("div", {"id":"df"})
        df_div.append(Soup(table, 'html.parser'))
        # Find and insert error table
        err_div = soup2.find("div", {"id":"error"})
        err_div.append(Soup(err_table, 'html.parser'))

        # write html file
        with open(html_file2,'w') as file:
            file.write(str(soup2))

        # check if file exists
        fName = ''.join([attach_file, '.pdf'])
        fnNum = 0

        while(os.path.isfile(fName) == True):
            fnNum += 1
            fName = ''.join([attach_file, '_', str(fnNum), '.pdf']) 

        # create PDF from html
        pdf_options = {'orientation': 'landscape',
                        'page-size': 'Letter',
                        'margin-top': '0.25in',
                        'margin-right': '0.25in',
                        'margin-bottom': '0.25in',
                        'margin-left': '0.25in',
                        'encoding': "UTF-8",}
                        
        pdfkit.from_file(html_file2, fName, options=pdf_options)

        # Email Claim Summary Report
        statusText = f'\r\nEmailing Claim Payment Summary report...'
        self.updateStatusText(statusText)

        to = ['jared@visualgap.com']
        sub = f'Claim Payment Summary {self.now.strftime("%Y-%m-%d")}'
        msg_html = '''
                    <html>
                    <body>
                        <p>Attached is the Claim Summary Report.<br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                '''
        if fnNum == 0:
            a_file = ''.join([attach_name, '.pdf'])
        else:
            a_file = ''.join([attach_name, '_', str(fnNum), '.pdf'])

        self.send_email(to, sub, msg_html, self.attachment_dir, a_file)

        # Create data file for SCC
        scc_cols = ['Carrier', 'claim_nbr_end', 'blank1', 'policy', 'blank2', 'last', 'first', 'loss_type', 'loss_date', 'blank3', 'GAP_static', 'GAP_amount', 'blank4', 
                    'blank5', 'blank6', 'lender', 'address', 'blank7', 'city', 'state', 'zip', 'status', 'blank8', 'GAP_static2', 'contract_id', 'vin', 'make',
                    'model', 'year', 'claim_nbr_front', 'CHECK', 'pymt_code', 'blank9', 'blank10', 'PAID_static', 'status_date', 'blank11', 'status_date2', 'blank12',
                    'FFS_static', 'chk_nbr']
        scc_df = pd.DataFrame(columns = scc_cols) 

        # Get SCC claims
        statusText = f'\r\nCreating SCC claim data file...'
        self.updateStatusText(statusText)
        scc_rpt_df = df[['rtbp_id', 'claim_id', 'amount', 'check_nbr', 'Claim_Type']].loc[(df['toVGC'] == 1) & (df['carrier_id'] == 8)].copy()

        # format amount without $, commas, and decimals
        scc_rpt_df.replace('[\$,\.]','', regex=True, inplace=True)

        # create scc_file df
        scc_file_df = pd.DataFrame()

        # Get SCC claim data
        if len(scc_rpt_df) > 0:
            for index, row in scc_rpt_df.iterrows():
                # Check or '' & payment code
                if row['amount'] == '$0.00':
                    check = ''
                    p_code = '003'
                else:
                    check = 'CHECK'
                    p_code = '001'
                # Paid or Plus
                if 'Plus' in row['Claim_Type']:
                    paid_or_plus = 'PLUS'
                else:
                    paid_or_plus = 'PAID'
                # ACH in chk_nbr to 0 
                if 'ACH' in row['check_nbr']:
                    chk_num = '0'
                else:
                    chk_num = row['check_nbr']            
                # create query
                sql = '''
                        SELECT 'SC' AS Carrier, 
                        SUBSTRING(c.claim_nbr, CHAR_LENGTH(c.claim_nbr)-5, 5) AS claim_nbr_end,
                        '' AS blank1,
                        l.policy_nbr,
                        '' AS blank2,
                        SUBSTRING(b.last, 1, 30) AS lastn,
                        SUBSTRING(b.first, 1, 30) AS firstn,
                        CASE
                            WHEN c.loss_type_id=1 THEN 'CO'
                            WHEN c.loss_type_id=2 THEN 'TH'
                            WHEN c.loss_type_id=3 THEN 'WE'
                            ELSE 'OT'
                        END AS loss_type,
                        DATE_FORMAT(c.loss_date, "%Y%m%d") AS loss_date,
                        '' AS blank3,
                        'GAP' AS GAP_static,
                        '{amt}' AS GAP_amount,
                        '' AS blank4,
                        '' AS blank5,
                        '' AS blank6,
                        SUBSTRING(l.alt_name, 1, 32) AS lender,
                        SUBSTRING(l.address1, 1, 32) AS address,
                        '' AS blank7,
                        l.city,
                        l.state,
                        SUBSTRING(l.zip, 1, 5) AS zip,
                        'C' AS status,
                        '' AS blank8,
                        'GAP' AS GAP_static2,
                        c.contractId,
                        v.vin,
                        SUBSTRING(v.make, 1, 20) AS Make,
                        SUBSTRING(v.model, 1, 20) AS Model,
                        SUBSTRING(v.year, 2, 2) AS Year,
                        SUBSTRING(c.claim_nbr, 1, 8) AS claim_nbr_front,
                        '{chk}' AS Chk,
                        '{pay_code}' AS pymt_code,
                        '' AS blank9,
                        '' AS blank10,
                        '{paid_plus}' AS PAID_static,
                        DATE_FORMAT(CURDATE(), "%Y%m%d") AS status_date,
                        '' AS blank11,
                        DATE_FORMAT(CURDATE(), "%Y%m%d") AS status_date2,
                        '' AS blank12,
                        'FFS' AS FFS_static,
                        '{check_nbr}' AS Chk_nbr       
                    FROM claims c
                    INNER JOIN claim_lender l
                        USING(claim_id)
                    INNER JOIN claim_borrower b
                        USING(claim_id)
                    INNER JOIN claim_vehicle v
                        USING(claim_id)
                    WHERE c.claim_id = {claimId};
                    '''.format(claimId=row['claim_id'], amt=row['amount'], chk=check, pay_code=p_code, paid_plus=paid_or_plus, check_nbr=chk_num)

                temp_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql, 0, 0))    
                scc_file_df = scc_file_df.append(temp_df)

        # fill each field has an exact starting position
        statusText = f'\r\nFormatting SCC claim data file...'
        self.updateStatusText(statusText)
        #A
        scc_file_df[0] = scc_file_df[0].str.pad(2, side='right', fillchar=' ')
        #B
        scc_file_df[1] = scc_file_df[1].str.pad(8, side='right', fillchar=' ')
        #C
        scc_file_df[2] = scc_file_df[2].str.pad(12, side='right', fillchar=' ')
        #D
        scc_file_df[3] = scc_file_df[3].str.pad(15, side='right', fillchar=' ')
        #E
        scc_file_df[4] = scc_file_df[4].str.pad(32, side='right', fillchar=' ')
        #F
        scc_file_df[5] = scc_file_df[5].str.pad(30, side='right', fillchar=' ')
        #G
        scc_file_df[6] = scc_file_df[6].str.pad(30, side='right', fillchar=' ')
        #H
        scc_file_df[7] = scc_file_df[7].str.pad(2, side='right', fillchar=' ')
        #I
        scc_file_df[8] = scc_file_df[8].str.pad(8, side='right', fillchar=' ')
        #J
        scc_file_df[9] = scc_file_df[9].str.pad(9, side='right', fillchar=' ')
        #K
        scc_file_df[10] = scc_file_df[10].str.pad(3, side='right', fillchar=' ')
        #L
        scc_file_df[11] = scc_file_df[11].str.pad(9, side='left', fillchar='0')
        #M
        scc_file_df[12] = scc_file_df[12].str.pad(9, side='right', fillchar=' ')
        #N
        scc_file_df[13] = scc_file_df[13].str.pad(9, side='right', fillchar=' ')
        #O
        scc_file_df[14] = scc_file_df[14].str.pad(10, side='right', fillchar=' ')
        #P
        scc_file_df[15] = scc_file_df[15].str.pad(32, side='right', fillchar=' ')
        #Q
        scc_file_df[16] = scc_file_df[16].str.pad(32, side='right', fillchar=' ')
        #R
        scc_file_df[17] = scc_file_df[17].str.pad(32, side='right', fillchar=' ')
        #S
        scc_file_df[18] = scc_file_df[18].str.pad(30, side='right', fillchar=' ')
        #T
        scc_file_df[19] = scc_file_df[19].str.pad(2, side='right', fillchar=' ')
        #U
        scc_file_df[20] = scc_file_df[20].str.pad(9, side='right', fillchar=' ')
        #V
        scc_file_df[21] = scc_file_df[21].str.pad(1, side='right', fillchar=' ')
        #W
        scc_file_df[22] = scc_file_df[22].str.pad(1, side='right', fillchar=' ')
        #X
        scc_file_df[23] = scc_file_df[23].str.pad(3, side='right', fillchar=' ')
        #Y
        scc_file_df[24] = scc_file_df[24].str.pad(20, side='right', fillchar=' ')
        #Z
        scc_file_df[25] = scc_file_df[25].str.pad(18, side='right', fillchar=' ')
        #AA
        scc_file_df[26] = scc_file_df[26].str.pad(20, side='right', fillchar=' ')
        #AB
        scc_file_df[27] = scc_file_df[27].str.pad(20, side='right', fillchar=' ')
        #AC
        scc_file_df[28] = scc_file_df[28].str.pad(2, side='right', fillchar=' ')
        #AD
        scc_file_df[29] = scc_file_df[29].str.pad(8, side='right', fillchar=' ')
        #AE
        scc_file_df[30] = scc_file_df[30].str.pad(20, side='right', fillchar=' ')
        #AF
        scc_file_df[31] = scc_file_df[31].str.pad(3, side='right', fillchar=' ')
        #AG
        scc_file_df[32] = scc_file_df[32].str.pad(20, side='right', fillchar=' ')
        #AH
        scc_file_df[33] = scc_file_df[33].str.pad(8, side='right', fillchar=' ')
        #AI
        scc_file_df[34] = scc_file_df[34].str.pad(20, side='right', fillchar=' ')
        #AJ
        scc_file_df[35] = scc_file_df[35].str.pad(8, side='right', fillchar=' ')
        #AK
        scc_file_df[36] = scc_file_df[36].str.pad(30, side='right', fillchar=' ')
        #AL
        scc_file_df[37] = scc_file_df[37].str.pad(8, side='right', fillchar=' ')
        #AM
        scc_file_df[38] = scc_file_df[38].str.pad(10, side='right', fillchar=' ')
        #AN
        scc_file_df[39] = scc_file_df[39].str.pad(3, side='right', fillchar=' ')
        #AO
        scc_file_df[40] = scc_file_df[40].str.pad(8, side='left', fillchar='0')

        # write text file
        statusText = f'\r\nSaving SCC claim data file...'
        self.updateStatusText(statusText)

        with open(self.attachment_dir + 'FrostGAP.txt', 'a') as f:
            for index, row in scc_file_df.iterrows():
                col_index = 0
                while col_index != len(row):
                    f.write(row[col_index])
                    col_index += 1
                f.write('\n')
        f.close()

        # Disable host key checking
        statusText = f'\r\nSending SCC claim data file via SFTP...'
        self.updateStatusText(statusText)

        cnopts = pysftp.CnOpts()
        cnopts.hostkeys = None

        # Send file to SCC via SFTP
        with pysftp.Connection(sftp_h, username=sftp_u, password=sftp_p, cnopts=cnopts) as sftp:
            with sftp.cd('dropoff'):
                sftp.put(self.attachment_dir + 'FrostGAP.txt')

        # Send Email notification to SCC and Claims
        to = ['jared@visualgap.com']
        sub = f'Frost claim file'
        msg_html = '''
                    <html>
                    <body>
                        <p>Hello,<br>
                        <br>
                        We have submitted a new claim file today.  If you have any questions or concerns please contact us. <br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                '''
        self.send_email(to, sub, msg_html)

        # update toVGC to 3
        statusText = f'\r\nUpdating payment status...'
        self.updateStatusText(statusText)

        if len(df) > 0:
            for index, row in df.iterrows():
                err_sql = '''
                        UPDATE ready_to_be_paid
                        SET toVGC = 3
                        WHERE rtbp_id = {rtbp_id};
                        '''.format(rtbp_id=row['rtbp_id'])
                # run update query        
                self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', err_sql, 0, 1)

        # Remove files from staging directory
        file_ext = [".csv", ".html"]
        # file_staging_dir = './letters/staging/'

        for ext in file_ext:
            self.clear_dir(self.file_staging_dir, ext)

        statusText = f'\r\nComplete!'
        self.updateStatusText(statusText)

    def vgcToQb(self):
        # Clear Output
        self.clearStatusText()

        # create batch_id
        # now = datetime.now()
        batch_id = self.now.strftime("%Y%m%d%H%M%S")

        statusText = f'Collecting GAP claim records...'
        self.updateStatusText(statusText)


        # sql query for GAP claims that are RTBP
        sql_file = '''
                    SELECT c.claim_id, c.claim_nbr, c.carrier_id, cl.alt_name, cl.dealer_securityId, cl.contact, cl.address1,
                        cl.city, cl.state, cl.zip, cl.payment_method, cb.first, cb.last, 
                        IF(sq.gap_amt_paid > 0, 2,1) AS pymt_type_id, 
                        IF(sq.gap_amt_paid > 0, ROUND(cc.gap_payable - sq.gap_amt_paid,2), cc.gap_payable) AS gap_due,
                        COALESCE(NULLIF(cb.acct_number,''),'0') AS acct_nbr, c.loss_date, cl.title, cl.email2
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

        # save query results as DF
        df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql_file, 0, 0))

        # add column names
        df_cols = ['claim_id', 'claim_nbr', 'carrier_id', 'lender_name', 'dealer_securityId', 'contact', 'address1', 'city', 'state', 'zip', 
                                    'pymt_method', 'first', 'last', 'pymt_type_id', 'amount', 'acct_number', 'loss_date', 'email', 'email2']
        df.columns = df_cols

        statusText = f'\r\nCollecting GAP Plus records...'
        self.updateStatusText(statusText)

        # sql query for PLUS claims that are RTBP
        sql_file_plus = '''
                        SELECT sqp.claim_id, c.claim_nbr, c.carrier_id, cl.alt_name, cl.dealer_securityId, cl.contact, 
                            cl.address1, cl.city, cl.state, cl.zip, cl.payment_method, cb.first, cb.last, 
                            1 AS pymt_type_id, 
                            1000 AS gap_plus_due, COALESCE(NULLIF(cb.acct_number,''),'0') AS acct_nbr,
                            c.loss_date, cl.title, cl.email2, cl.customer_securityId
                        FROM claims c
                        INNER JOIN claim_lender cl
                            USING (claim_id)
                        INNER JOIN claim_borrower cb
                            USING (claim_id)
                        INNER JOIN (SELECT pb.claim_id
                                    FROM claim_plus_benefit pb
                                    WHERE status_desc_id = 8) sqp
                            USING (claim_id)
                        INNER JOIN (SELECT c.claim_id
                                    FROM claims c
                                    INNER JOIN claim_status cs
                                        ON (c.status_id = cs.status_id)  
                                    WHERE cs.status_desc_id = 8
                                        OR cs.status_desc_id = 4) sqg
                            USING (claim_id);
                    '''
        
        # save query results as DF
        df2 = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql_file_plus, 0, 0))

        # add column names
        df2_cols = ['claim_id', 'claim_nbr', 'carrier_id', 'lender_name', 'dealer_securityId', 'contact', 'address1', 'city', 'state', 'zip', 
                                    'pymt_method', 'first', 'last', 'pymt_type_id', 'amount', 'acct_number', 'loss_date', 'email', 'email2', 'CUCode']
        df2.columns = df2_cols

        # GAP Plus exceptions (benefits other $1,000)
        # Tulare (9401) -  $1,500
        # West AirComm (355) - $2,500

        # TEMPORARY ########################################################################################
        # Convert test customer_securityId to Production customer_securityId
        df2['CUCode'].replace({18475:9401,20828:355}, inplace=True)
        # TEMPORARY ########################################################################################

        # update benefit amount
        df2.loc[(df2.CUCode == 9401),'amount']=1500
        df2.loc[(df2.CUCode == 355),'amount']=2500

        # drop CUCode column
        df2.drop(columns=['CUCode'], inplace=True)

        statusText = f'\r\nCollecting TotalRestart records...'
        self.updateStatusText(statusText)

        # sql query for TotalRestart claims that are RTBP
        # manually entered carrier_id 12
        sql_file_tr = '''
                    SELECT sqp.claim_id, c.claim_nbr, 12 AS carrier_id, cl.alt_name, cl.dealer_securityId, cl.contact, cl.address1, 
                    cl.city, cl.state, cl.zip, 'Check' AS payment_method, cb.first, cb.last, 1 AS pymt_type_id, 
                    ctr.totalrestart_payable AS tr_due, COALESCE(NULLIF(cb.acct_number,''),'0') AS acct_nbr, c.loss_date, cl.title, cl.email2
                    FROM claims c
                    INNER JOIN claim_lender cl
                        USING (claim_id)
                    INNER JOIN claim_borrower cb
                        USING (claim_id)
                    INNER JOIN (SELECT pb.claim_id
                                FROM claim_totalrestart pb
                                WHERE status_desc_id = 8) sqp
                        USING (claim_id)
                    INNER JOIN claim_totalrestart ctr
                        USING (claim_id)
                    '''
        
        # save query results as DF
        df3 = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql_file_tr, 0, 0))

        # add column names
        df3_cols = ['claim_id', 'claim_nbr', 'carrier_id', 'lender_name', 'dealer_securityId', 'contact', 'address1', 'city', 'state', 'zip', 
                                    'pymt_method', 'first', 'last', 'pymt_type_id', 'amount', 'acct_number', 'loss_date', 'email', 'email2']
        df3.columns = df3_cols

        # GAP
        # convert columns list to string
        cols = ", ".join(df_cols)

        # insert DF into the ready_to_be_paid table of the claim_qb_payments database
        for x,rows in df.iterrows():

            sql_file2 = '''INSERT INTO ready_to_be_paid ({columns}, payment_category_id, check_nbr, batch_id, qb_txnid, toVGC) VALUES ({claim_id}, "{claim_nbr}", {carrier_id},"{lender_name}", "{lender_id}","{contact}", 
                        "{address1}", "{city}", "{state}", "{zip}", "{pymt_method}", "{first}", "{last}", {pymt_type_id}, {amount}, "{acct_nbr}", "{loss_date}", "{email}", "{email2}", 1, 0, {batchId}, 0, 0);'''.format(columns=cols, 
                        claim_id=rows['claim_id'], claim_nbr=rows['claim_nbr'], carrier_id=rows['carrier_id'], lender_name=rows['lender_name'], lender_id=rows['dealer_securityId'], 
                        contact=rows['contact'], address1=rows['address1'], city=rows['city'], state=rows['state'], zip=rows['zip'], 
                        pymt_method=rows['pymt_method'], first = rows['first'], last = rows['last'], pymt_type_id = rows['pymt_type_id'], 
                        amount = rows['amount'], acct_nbr = rows['acct_number'], loss_date = rows['loss_date'], email= rows['email'], email2= rows['email2'], batchId = batch_id)
            
            # # execute and commit sql
            self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql_file2, 0, 1)

        # PLUS
        # convert columns list to string
        cols = ", ".join(df3_cols)

        # insert DF into the ready_to_be_paid table of the claim_qb_payments database
        for x,rows in df2.iterrows():
            sql_file2_plus = '''INSERT INTO ready_to_be_paid ({columns}, payment_category_id, check_nbr, batch_id, qb_txnid, toVGC) VALUES ({claim_id}, "{claim_nbr}", {carrier_id},"{lender_name}", "{lender_id}","{contact}", 
                        "{address1}", "{city}", "{state}", "{zip}", "{pymt_method}", "{first}", "{last}", {pymt_type_id}, {amount}, "{acct_nbr}", "{loss_date}", "{email}", "{email2}", 2, 0, {batchId}, 0, 0);'''.format(columns=cols, 
                        claim_id=rows['claim_id'], claim_nbr=rows['claim_nbr'], carrier_id=rows['carrier_id'], lender_name=rows['lender_name'], lender_id=rows['dealer_securityId'], 
                        contact=rows['contact'], address1=rows['address1'], city=rows['city'], state=rows['state'], zip=rows['zip'], 
                        pymt_method=rows['pymt_method'], first = rows['first'], last = rows['last'], pymt_type_id = rows['pymt_type_id'], 
                        amount = rows['amount'], acct_nbr = rows['acct_number'], loss_date = rows['loss_date'], email= rows['email'], email2= rows['email2'], batchId = batch_id)
                
            # execute and commit sql
            self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql_file2_plus, 0, 1)

        # TOTALRESTART
        # convert columns list to string
        cols = ", ".join(df3_cols)

        # insert DF into the ready_to_be_paid table of the claim_qb_payments database
        for x,rows in df3.iterrows():
            sql_file2_tr = '''INSERT INTO ready_to_be_paid ({columns}, payment_category_id, check_nbr, batch_id, qb_txnid, toVGC) VALUES ({claim_id}, "{claim_nbr}", {carrier_id},"{lender_name}", "{lender_id}","{contact}", 
                        "{address1}", "{city}", "{state}", "{zip}", "{pymt_method}", "{first}", "{last}", {pymt_type_id}, {amount}, "{acct_nbr}", "{loss_date}", "{email}", "{email2}", 3, 0, {batchId}, 0, 0);'''.format(columns=cols, 
                        claim_id=rows['claim_id'], claim_nbr=rows['claim_nbr'], carrier_id=rows['carrier_id'], lender_name=rows['lender_name'], lender_id=rows['dealer_securityId'], 
                        contact=rows['contact'], address1=rows['address1'], city=rows['city'], state=rows['state'], zip=rows['zip'], 
                        pymt_method=rows['pymt_method'], first = rows['first'], last = rows['last'], pymt_type_id = rows['pymt_type_id'], 
                        amount = rows['amount'], acct_nbr = rows['acct_number'], loss_date = rows['loss_date'], email= rows['email'], email2= rows['email2'], batchId = batch_id)
                
            # execute and commit sql
            self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql_file2_tr, 0, 1)

        # create select query to pull current batch with ID
        sql_file3 = '''SELECT rtbp_id, claim_id, claim_nbr, carrier_id, lender_name, dealer_securityId, contact, address1, city, 
                            state, zip, pymt_method, first, last, pymt_type_id, amount, payment_category_id, check_nbr, batch_id, 
                            qb_txnid, acct_number, loss_date, email, email2
                    FROM ready_to_be_paid
                    WHERE batch_id = {batchId}
                    ORDER BY payment_category_id, claim_id;'''.format(batchId = batch_id)

        # save query results as DF
        pymts_df = pd.DataFrame(self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql_file3, 0, 0))

        # add column names
        pymts_df_cols = ['rtbp_id', 'claim_id', 'claim_nbr', 'carrier_id', 'lender_name', 'dealer_securityId', 'contact', 'address1', 'city', 'state', 'zip', 
                            'pymt_method', 'first', 'last', 'pymt_type_id', 'amount','payment_category_id', 'check_nbr', 'batch_id', 'qb_txnid', 
                            'acct_number', 'loss_date', 'email', 'email2']
        pymts_df.columns = pymts_df_cols

        statusText = f'\r\nCollecting QB List IDs...'
        self.updateStatusText(statusText)

        # sql query for expense accounts in DB
        sql_file4 = '''
            SELECT carrier_id, qb_fullname
            FROM qb_accounts
            WHERE account_type = 'Expense';
            '''

        # save query results as DF
        expense_df = pd.DataFrame(self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql_file4, 0, 0))

        # add column names to DF
        col_names = ['carrier_id', 'expense']
        expense_df.columns = col_names

        # sql query for checking accounts in DB
        sql_file5 = '''
            SELECT carrier_id, qb_fullname
            FROM qb_accounts
            WHERE account_type = 'Checking';
            '''

        # save query results as DF
        checking_df = pd.DataFrame(self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', sql_file5, 0, 0))

        # add column names to DF
        col_names = ['carrier_id', 'checking']
        checking_df.columns = col_names

        # Merge expense account name into df
        pymts_df = pymts_df.merge(expense_df, how='left', left_on='carrier_id', right_on='carrier_id').copy()

        # Merge checking account name into df
        pymts_df = pymts_df.merge(checking_df, how='left', left_on='carrier_id', right_on='carrier_id').copy()

        statusText = f'\r\nChecking records for errors...'
        self.updateStatusText(statusText)

        # Create Error DF
        error_df = pd.DataFrame(columns = ['rtbp_id', 'err_msg'])

        # Check for missing expense and/or checking names
        chk_exp_error_df = pymts_df.loc[(pymts_df['expense'].isnull()) | (pymts_df['checking'].isnull())]

        # Drop rows with error
        pymts_df.drop(pymts_df[(pymts_df['expense'].isnull()) | (pymts_df['checking'].isnull())].index, inplace = True)

        # add to Error DF
        if len(chk_exp_error_df) > 0:
            for index, row in chk_exp_error_df.iterrows():
                error_df.loc[error_df.shape[0]] = [row['rtbp_id'], 'No matching CHECKING and-or EXPENSE accounts in QuickBooks'] 

        # Create sql server connection
        cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+svr+';DATABASE='+db+';UID='+sql_u+';PWD='+ sql_pw)

        # create query
        sql_svr_file = '''SELECT VGSecurityId, QB_ListID
                        FROM business_entity
                        WHERE QB_ListID IS NOT NULL;
                    '''
        # execute query
        listid_df = pd.read_sql(sql_svr_file, cnxn)

        # TEMPORARY ########################################################################################
        # Convert test dealer_securityId to Production dealer_securityId
        pymts_df['dealer_securityId'].replace({22260:46724,21945:52715,21095:9401,21519:355}, inplace=True)
        # TEMPORARY ########################################################################################

        # Merge QB_ListID into df
        pymts_df = pymts_df.merge(listid_df, how='left', left_on='dealer_securityId', right_on='VGSecurityId').copy()

        # check claims for Policy and Contract ID
        scc_df = pymts_df[['claim_id', 'claim_nbr']].loc[pymts_df['carrier_id'] == 8].copy()
        scc_df.set_index('claim_id', inplace=True)
        scc_df[['contractId', 'policy_nbr']] = ['','']

        statusText = f'\r\nCollect Policy Number and Contract ID from VGC...'
        self.updateStatusText(statusText)

        # add policy number and contract id to df
        if len(scc_df) > 0:
            for index, row in scc_df.iterrows():
                sql = '''
                    SELECT c.claim_id, c.contractId, l.policy_nbr 
                    FROM claims c
                    INNER JOIN claim_lender l
                        USING(claim_id)
                    WHERE c.claim_id = {claimId};
                    '''.format(claimId = index)
                
                # run sql query
                scc_p_c_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql, 0, 0))

                cols = ['claim_id', 'contractId', 'policy_nbr']
                scc_p_c_df.columns = cols
                scc_p_c_df.set_index('claim_id', inplace=True)

                # Update scc_df with scc_p_c_df
                scc_df.update(scc_p_c_df)

        statusText = f'\r\nCheck for missing Policy Number and Contract IDs...'
        self.updateStatusText(statusText)

        # check for missing policy number and/or contract id for SCC claims
        missing_p_c_df = scc_df.loc[(scc_df['contractId'] == '') | (scc_df['policy_nbr'] == '')]
        missing_p_c_df.set_index('claim_nbr', inplace=True)

        # if any claims are missing policy and contract id, 1) create a report, 2) email the report & 3) exit script.
        if len(missing_p_c_df) > 0:
            html_file = './letters/staging/missing_data.html'
            csv_file = './letters/staging/missing_data.csv'
            template_file = './html/missing_data_template.html'
            html_file2 = './letters/staging/missing_data.html'
            # export df as csv
            missing_p_c_df.to_csv(csv_file)
            # read in csv
            csvFile = pd.read_csv(csv_file)
            # convert csv to html
            csvFile.to_html(html_file, index=False)
            # Get html df
            soup=Soup(open(html_file), "html.parser")
            table = str(soup.select_one("table", {"class":"dataframe"}))
            # Get template
            soup2 = Soup(open(template_file), "html.parser")
            # Find and insert payment table
            df_div = soup2.find("div", {"id":"df"})
            df_div.append(Soup(table, 'html.parser'))
            
            # Email Claim Summary Report
            to = ['jared@visualgap.com']
            sub = f'Claims Missing Data'

            self.send_email(to, sub, soup2)

            #exit script
            statusText = f'\r\nThere are missing Policy Number(s) and/or Contract Id(s) in VGC\r\nAn email was sent with details\r\nProcess Cancelled.'

            self.updateStatusText(statusText)
            sys.exit("Missing data from SCC claims.")

        # Check for missing QBlistID
        qbListId_error_df = pymts_df.loc[(pymts_df['QB_ListID'].isnull())]

        statusText = f'\r\nChecking records for errors...'
        self.updateStatusText(statusText)

        # add to Error DF
        if len(qbListId_error_df) > 0:
            for index, row in qbListId_error_df.iterrows():
                error_df.loc[error_df.shape[0]] = [row['rtbp_id'], 'Missing QB_LIST_ID in the Claim_Payments database.  Run GET CUSTOMER to update the database.']

        # Using error_df, send error messages to ready_to_be_paid DB and update toVGC = 2
        if len(error_df) > 0:
            for index, row in error_df.iterrows():
                err_sql = '''
                        UPDATE ready_to_be_paid
                        SET toVGC = 2,
                            err_msg = '{err_msg}'
                        WHERE rtbp_id = {rtbp_id};
                        '''.format(err_msg=row['err_msg'], rtbp_id=row['rtbp_id'])

                self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', err_sql, 0, 1)

        # Drop rows with error
        pymts_df.drop(pymts_df[pymts_df['QB_ListID'].isnull()].index, inplace = True)

        # Add carrier name
        sql_file6 = '''
            SELECT carrier_id, description
            FROM carriers;
            '''

        # save query results as DF
        carrier_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql_file6, 0, 0))

        col_names = ['carrier_id', 'carrier']
        carrier_df.columns = col_names

        # Merge QB_ListID into df
        pymts_df = pymts_df.merge(carrier_df, left_on='carrier_id', right_on='carrier_id').copy()

        # payments greater than 0 df
        qb_pymts_df = pymts_df.loc[pymts_df['amount'] > 0].copy()

        statusText = f'\r\nConnecting to QuickBooks...'
        self.updateStatusText(statusText)

        # Connect to Quickbooks
        try:
            sessionManager = wc.Dispatch("QBXMLRP2.RequestProcessor")    
            sessionManager.OpenConnection('', 'Claim Payments')
            ticket = sessionManager.BeginSession("", 2)
        except Exception as e:
            print('''
            Make sure QuickBooks is running and you are logged into the Company File.
            ERROR: {}'''.format(e))

            # create and record error messages in ready_to_be_paid DB and update toVGC = 2
            for index, row in pymts_df.iterrows():
                error_df.loc[error_df.shape[0]] = [row['rtbp_id'], 'Unable to connect to QuickBooks']

            if len(error_df) > 0:
                for index, row in error_df.iterrows():
                    err_sql = '''
                            UPDATE ready_to_be_paid
                            SET toVGC = 2,
                                err_msg = '{err_msg}'
                            WHERE rtbp_id = {rtbp_id};
                            '''.format(err_msg=row['err_msg'], rtbp_id=row['rtbp_id'])

                    self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', err_sql, 0, 1)

            statusText = f'\r\nThere was an ERROR connecting to QuickBooks.\r\n***Make sure QuickBooks desktop is open and you are logged into the correct Company file.***\r\nProcess Cancelled.'
            self.updateStatusText(statusText)

            return

        # create qbxml query to add payments to QB
        pay_date = f"{dt.date.today():%Y-%m-%d}"

        statusText = f'\r\nRunning query...'
        self.updateStatusText(statusText)

        for index, row in qb_pymts_df.iterrows():
            if row['payment_category_id'] == 1:
                if row['pymt_type_id'] == 2:
                    pymt_type = 'Additional GAP Claim Pymt'
                else: 
                    pymt_type = 'GAP Claim'

            elif row['payment_category_id'] == 2:
                if row['pymt_type_id'] == 2:
                    pymt_type = 'Additional GAP Plus Pymt'
                else: 
                    pymt_type = 'GAP Plus'
                    
            elif row['payment_category_id'] == 3:
                if row['pymt_type_id'] == 2:
                    pymt_type = 'Additional TotalRestart Pymt'
                else: 
                    pymt_type = 'TotalRestart'

            pymtAmt = "{:.2f}".format(row['amount'])

            if row['pymt_method'] == 'Check':
                qbxmlQuery = '''
                <?qbxml version="14.0"?>
                <QBXML>
                    <QBXMLMsgsRq onError="stopOnError">
                        <CheckAddRq>
                            <CheckAdd>
                                <AccountRef>
                                    <FullName>{checking}</FullName>
                                </AccountRef>
                                <PayeeEntityRef>
                                    <ListID>{lender_qbid}</ListID>
                                </PayeeEntityRef>
                                <TxnDate>{date}</TxnDate>
                                <Memo>{memo}</Memo>
                                <Address>
                                    <Addr1>{lender}</Addr1>
                                    <Addr2>{contact}</Addr2>
                                    <Addr3>{address}</Addr3>
                                    <City>{city}</City>
                                    <State>{state}</State>
                                    <PostalCode>{zip}</PostalCode>
                                </Address>
                                <IsToBePrinted>true</IsToBePrinted>
                                <ExpenseLineAdd>
                                    <AccountRef>
                                        <FullName>{expense}</FullName>
                                    </AccountRef>
                                    <Amount>{amount}</Amount>
                                    <Memo>{memo}</Memo>
                                </ExpenseLineAdd>
                            </CheckAdd>
                        </CheckAddRq>
                    </QBXMLMsgsRq>
                </QBXML>'''.format(checking=row['checking'], lender_qbid=row['QB_ListID'], lender=row['lender_name'], date=pay_date, 
                        memo=f"{row['last']}/{row['first']} {pymt_type}", contact=row['contact'], address=row['address1'], city=row['city'], state=row['state'],
                        zip=row['zip'], expense=row['expense'], amount = pymtAmt)

            elif 'ACH' in row['pymt_method']:
                qbxmlQuery = '''
                <?qbxml version="14.0"?>
                <QBXML>
                    <QBXMLMsgsRq onError="stopOnError">
                        <CheckAddRq>
                            <CheckAdd>
                                <AccountRef>
                                    <FullName>{checking}</FullName>
                                </AccountRef>
                                <PayeeEntityRef>
                                    <ListID>{lender_qbid}</ListID>
                                </PayeeEntityRef>
                                <RefNumber>ACH</RefNumber>
                                <TxnDate>{date}</TxnDate>
                                <Memo>{memo}</Memo>
                                <Address>
                                    <Addr1>{lender}</Addr1>
                                    <Addr2>{contact}</Addr2>
                                    <Addr3>{address}</Addr3>
                                    <City>{city}</City>
                                    <State>{state}</State>
                                    <PostalCode>{zip}</PostalCode>
                                </Address>
                                <IsToBePrinted>false</IsToBePrinted>
                                <ExpenseLineAdd>
                                    <AccountRef>
                                        <FullName>{expense}</FullName>
                                    </AccountRef>
                                    <Amount>{amount}</Amount>
                                    <Memo>{memo}</Memo>
                                </ExpenseLineAdd>
                            </CheckAdd>
                        </CheckAddRq>            
                    </QBXMLMsgsRq>
                </QBXML>'''.format(checking=row['checking'], lender_qbid=row['QB_ListID'], lender=row['lender_name'], date=pay_date, 
                        memo=f"{row['last']}/{row['first']} {pymt_type}", contact=row['contact'], address=row['address1'], city=row['city'], state=row['state'],
                        zip=row['zip'], expense=row['expense'], amount = pymtAmt)

            # Send query and receive response
            responseString = sessionManager.ProcessRequest(ticket, qbxmlQuery)

            # output TxnID
            QBXML = ET.fromstring(responseString)
            QBXMLMsgsRs = QBXML.find('QBXMLMsgsRs')
            checkResults = QBXMLMsgsRs.iter("CheckRet")
            txnId = 0
            for checkResult in checkResults:
                txnId = checkResult.find('TxnID').text

            # Add TxnID to ready_to_be_paid table
            qb_sql_file = '''UPDATE ready_to_be_paid
                        SET qb_txnid = '{TxnID}',
                            pymt_date = '{paydate}'
                        WHERE rtbp_id = {rowID};'''.format(TxnID=txnId, paydate=pay_date, rowID=row['rtbp_id'])
            
            # execute and commit sql
            self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', qb_sql_file, 0, 1)

        # Disconnect from Quickbooks
        statusText = f'\r\nDisconnecting from Quickbooks...'
        self.updateStatusText(statusText)

        sessionManager.EndSession(ticket)
        sessionManager.CloseConnection()

        # Add Paid Date to $0 claims
        zero_pymt_df = pymts_df.loc[pymts_df['amount'] == 0].copy()
        for index, row in zero_pymt_df.iterrows():
            #sql query
            zero_sql_file = '''UPDATE ready_to_be_paid
                        SET pymt_date = '{paydate}'
                        WHERE rtbp_id = {rowID};'''.format(paydate=pay_date, rowID=row['rtbp_id'])
            
            # execute and commit sql
            self.mysql_q(mysql_u, mysql_pw, mysql_host, 'claim_qb_payments', zero_sql_file, 0, 1)

        statusText = f'\r\nCollecting Fraud Language...'
        self.updateStatusText(statusText)

        # Add Fraud Language fields to pymts_df
        sql_file7 = '''
            SELECT StateId, 
            StateDesc,
            StateCode,
            CAST(Language AS CHAR(1000) CHARACTER SET utf8)
            FROM FraudLang;
            '''

        # run sql query
        fraud_lang_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', sql_file7, 0, 0))

        col_names = ['StateId', 'StateDesc', 'StateCode', 'f_lang']
        fraud_lang_df.columns = col_names

        # Merge QB_ListID into df
        pymts_df = pymts_df.merge(fraud_lang_df, how='left', left_on='state', right_on='StateId').copy()

        # replace NaN with ''
        pymts_df.fillna('', inplace = True)

        # Update Column names
        pymts_df.rename(columns = {'amount':'payment_amount', 'lender_name':'alt_name'}, inplace = True)

        ##### CREATE GAP LETTERS #####

        statusText = f'\r\nCreating GAP Letters and Calculations...'
        self.updateStatusText(statusText)

        # Collect GAP payments
        gap_pymts_df = pymts_df.loc[pymts_df['payment_category_id'] == 1].copy()
        gap_letters_df = gap_pymts_df.loc[gap_pymts_df['payment_amount'] > 0].copy()

        # Remove files from staging directory
        # file_staging_dir = './letters/staging/'
        file_ext = ".pdf"

        self.clear_dir(self.file_staging_dir, file_ext)

        ##### create letters for amounts greater than $0 paid via check #####

        # Filter gap_letters_df for pymt_method 'Check'
        checks_df = gap_letters_df.loc[gap_letters_df['pymt_method'] == 'Check'].copy()

        if len(checks_df.index) > 0:
            # Create list of Carriers on checks_df
            carriers = []
            carriers = checks_df.carrier.unique()
            letter_cols = ['claim_nbr', 'loss_date', 'alt_name', 'contact', 'address1', 'city', 'state', 'zip', 'first', 'last', 'acct_number', 'payment_amount', 'StateDesc', 'StateCode', 'f_lang' ]

            # Loop through carriers list
            for carrier in carriers:

                # Variable Defaults
                sql_where_cal = ''
                g_letters = True
                s_letters = True
                cals = []

                # check for payment_type_id for GAP letters by carrier
                g_letters_df = checks_df.loc[(checks_df['pymt_type_id'] == 1) & (checks_df['carrier'] == carrier)].copy()

                # GAP Letter - create WHERE statement
                if len(g_letters_df.index) > 0:
                    g_letters = True
                else:
                    # No Letters
                    g_letters = False

                # check for payment_type_id for Supplemental letters
                s_letters_df = checks_df.loc[(checks_df['pymt_type_id'] == 2) & (checks_df['carrier'] == carrier)].copy()

                # Supplemental Letter
                if len(s_letters_df.index) > 0:
                    s_letters = True
                else:
                    # No Letters
                    s_letters = False

                # Calculations
                # check for calculations for carrier
                calculations_df = checks_df.loc[checks_df['carrier'] == carrier].copy()

                if len(calculations_df.index) > 0:
                    # Multiple Calculation
                    for index, row in calculations_df.iterrows():
                        cals.append(row['claim_id'])

                # GAP - Create letters
                if g_letters == True:
                    # create df for template
                    g_template_df = g_letters_df[letter_cols]

                    # Create GAP Letters
                    g_pdf_template = "letters/pdf_templates/GAP_letter_template.pdf"
                    position = 1
                    self.gap_letter(g_template_df, g_pdf_template, position)
                
                # Supplemental - Create letters
                if s_letters == True:
                    # create df for template
                    s_template_df = s_letters_df[letter_cols]
                    
                    # Create Supplement Letters
                    s_pdf_template = "letters/pdf_templates/Supp_letter_template.pdf"
                    position = 2
                    self.gap_letter(s_template_df, s_pdf_template, position)

                # Calculations - Create SQL query, run query, create calculation sheets
                if g_letters == True or s_letters == True:

                    # create calculation dataframe
                    temp_cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42]                             
                    c_consolidate_df = pd.DataFrame(columns = temp_cols)

                    for cal in cals:
                        # create sql for calculation
                        c_sql_query = '''
                                SELECT c.claim_nbr, c.loss_date, l.alt_name, l.contact, b.first, b.last, cc.gap_payable,
                                    cl.incp_date, cl.last_payment, cl.interest_rate, cl.amount AS Amt_Fin, cc.balance_last_pay,
                                    cc.nbr_of_days, cc.per_day, ROUND(cc.payoff,2) AS payoff, (cc.ltv * 100) AS ltv, cc.covered_fin_amount,
                                    (cc.percent_uncovered * 100) AS percent_uncovered, (cc.ltv_limit * 100) AS ltv_limit, v.nada_value, CONCAT(v.year, ' ', v.make, ' ', v.model) AS vehicle,
                                    v.deductible, cls.description AS loss_type, cc26.description AS primary_carrier,
                                    cc14.amount AS past_due, cc15.amount AS late_fees, cc16.amount AS skip_pymts,
                                    cc17.amount AS skip_fees, cc5.amount AS primary_pymt, cc7.amount AS excess_deductible,
                                    cc8.amount AS scr, cc9.amount AS clr, cc10.amount AS cdr, cc11.amount AS oref,
                                    cc18.amount AS salvage, cc19.amount AS prior_dmg, cc20.amount AS over_ltv,
                                    cc21.amount AS other1_amt, cc21.description AS other1_description,
                                    cc22.amount AS other2_amt, cc22.description AS other2_description, ca.description AS carrier,
                                    (cc22.amount + cc21.amount + cc20.amount + cc19.amount + cc18.amount + cc11.amount + cc10.amount + cc9.amount + 
                                    cc8.amount + cc7.amount + cc5.amount + cc17.amount + cc16.amount + cc15.amount + cc14.amount) as subtotal
                                FROM claims c
                                INNER JOIN claim_lender l
                                    USING (claim_id)
                                INNER JOIN claim_borrower b
                                    USING (claim_id)
                                INNER JOIN claim_loan cl
                                    USING (claim_id)
                                INNER JOIN claim_calculations cc
                                    USING (claim_id)
                                INNER JOIN claim_vehicle v
                                    USING (claim_id)
                                INNER JOIN claims_loss_type cls
                                    USING (loss_type_id)
                                INNER JOIN carriers ca
                                    ON c.carrier_id = ca.carrier_id                        
                                INNER JOIN claim_checklist AS cc26
                                    ON c.claim_id = cc26.claim_id
                                    AND cc26.checklist_item_id = 26    
                                INNER JOIN claim_checklist AS cc14
                                    ON c.claim_id = cc14.claim_id
                                    AND cc14.checklist_item_id = 14    
                                INNER JOIN claim_checklist AS cc15
                                    ON c.claim_id = cc15.claim_id
                                    AND cc15.checklist_item_id = 15    
                                INNER JOIN claim_checklist AS cc16
                                    ON c.claim_id = cc16.claim_id
                                    AND cc16.checklist_item_id = 16       
                                INNER JOIN claim_checklist AS cc17
                                    ON c.claim_id = cc17.claim_id
                                    AND cc17.checklist_item_id = 17       
                                INNER JOIN claim_checklist AS cc5
                                    ON c.claim_id = cc5.claim_id
                                    AND cc5.checklist_item_id = 5
                                INNER JOIN claim_checklist AS cc7
                                    ON c.claim_id = cc7.claim_id
                                    AND cc7.checklist_item_id = 7
                                INNER JOIN claim_checklist AS cc8
                                    ON c.claim_id = cc8.claim_id
                                    AND cc8.checklist_item_id = 8
                                INNER JOIN claim_checklist AS cc9
                                    ON c.claim_id = cc9.claim_id
                                    AND cc9.checklist_item_id = 9
                                INNER JOIN claim_checklist AS cc10
                                    ON c.claim_id = cc10.claim_id
                                    AND cc10.checklist_item_id = 10
                                INNER JOIN claim_checklist AS cc11
                                    ON c.claim_id = cc11.claim_id
                                    AND cc11.checklist_item_id = 11      
                                INNER JOIN claim_checklist AS cc18
                                    ON c.claim_id = cc18.claim_id
                                    AND cc18.checklist_item_id = 18
                                INNER JOIN claim_checklist AS cc19
                                    ON c.claim_id = cc19.claim_id
                                    AND cc19.checklist_item_id = 19
                                INNER JOIN claim_checklist AS cc20
                                    ON c.claim_id = cc20.claim_id
                                    AND cc20.checklist_item_id = 20     
                                INNER JOIN claim_checklist AS cc21
                                    ON c.claim_id = cc21.claim_id
                                    AND cc21.checklist_item_id = 21
                                INNER JOIN claim_checklist AS cc22
                                    ON c.claim_id = cc22.claim_id
                                    AND cc22.checklist_item_id = 22  
                                WHERE c.claim_id = {claimID};'''.format(claimID = cal)

                        # save query results as DF
                        c_temp_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', c_sql_query, 0, 0))

                        # append query result to c_template_df
                        c_consolidate_df = c_consolidate_df.append(c_temp_df)

                    cal_cols = {0:'claim_nbr', 1:'loss_date', 2:'alt_name', 3:'contact', 4:'first', 5:'last', 6:'gap_payable', 7:'incp_date', 8:'last_payment', 9:'interest_rate', 10:'Amt_Fin',
                        11:'balance_last_pay', 12:'nbr_of_days', 13:'per_day', 14:'payoff', 15:'ltv', 16:'covered_fin_amount', 17:'percent_uncovered', 18:'ltv_limit', 19:'nada_value',
                        20:'vehicle', 21:'deductible', 22:'loss_type', 23:'primary_carrier', 24:'past_due', 25:'late_fees', 26:'skip_pymts', 27:'skip_fees', 28:'primary_pymt',
                        29:'excess_deductible', 30:'scr', 31:'clr', 32:'cdr', 33:'oref', 34:'salvage', 35:'prior_dmg', 36:'over_ltv', 37:'other1_amt', 38:'other1_description',
                        39:'other2_amt', 40:'other2_description', 41:'carrier', 42:'subtotal'}

                    c_template_df = c_consolidate_df.rename(columns = cal_cols)
                    c_template_df = c_template_df.reset_index(drop=True)

                    # Create Calculations
                    c_pdf_template = "letters/pdf_templates/GAP_calculation_template.pdf"
                    position = 3
                    self.calculations(c_template_df, c_pdf_template, position)
                        
                # Close SQL Connection

                # create a list of file to concatenate
                file_list = self.fileList(self.file_staging_dir, file_ext)

                # Concatenated output file
                outfn = f'S:/claims/letters/{carrier}_{self.now.strftime("%Y-%m-%d")}_GAP'

                self.ConCat_pdf(file_list, outfn)

                # Remove files from staging directory
                self.clear_dir(self.file_staging_dir, file_ext)

        else: print('No amount greater then 0.')

        # Update toVGC to 1
        self.update_tovgc_1(checks_df)

        ##### CREATE PLUS LETTERS #####

        statusText = f'\r\nCreating GAP Plus Letters...'
        self.updateStatusText(statusText)

        # Collect GAP Plus
        plus_pymts_df = pymts_df.loc[pymts_df['payment_category_id'] == 2].copy()
        plus_letters_df = plus_pymts_df.loc[plus_pymts_df['payment_amount'] > 0].copy()

        # Filter plus_letters_df for pymt_method 'Check'
        plus_df = plus_letters_df.loc[plus_letters_df['pymt_method'] == 'Check'].copy()

        if len(plus_df.index) > 0:

            # Create list of Carriers on checks_df
            p_carriers = []
            p_carriers = plus_df.carrier.unique()

            # Loop through carriers list
            for p_carrier in p_carriers:

                # Variable Defaults
                p_letters = True

                # check for payment_type_id for GAP letters by carrier
                p_letters_df = plus_df.loc[plus_df['carrier'] == p_carrier].copy()

                # check for records
                if len(p_letters_df.index) > 0:
                    p_letters = True
                else:
                    # No Letters
                    p_letters = False

                # PLUS - Create letters
                if p_letters == True:
                    # create df for template
                    p_template_df = p_letters_df[letter_cols]

                    # Create GAP Letters
                    p_pdf_template = "letters/pdf_templates/PLUS_letter_template.pdf"
                    position = 1
                    self.gap_letter(p_template_df, p_pdf_template, position)
                
                    file_list = self.fileList(self.file_staging_dir, file_ext)

                    # Concatenated output file
                    outfn = f'S:/claims/letters/{p_carrier}_{self.now.strftime("%Y-%m-%d")}_PLUS'

                    self.ConCat_pdf(file_list, outfn)

                # Remove files from staging directory
                self.clear_dir(self.file_staging_dir, file_ext)

        else: print('No Plus.')

        # Update toVGC to 1
        self.update_tovgc_1(plus_df)

        ##### CREATE TOTALRESTART LETTERS #####

        statusText = f'\r\nCreating TotalRestart Letters and Calculations...'
        self.updateStatusText(statusText)

        # Collect GAP Plus
        tr_pymts_df = pymts_df.loc[pymts_df['payment_category_id'] == 3].copy()
        tr_letters_df = tr_pymts_df.loc[tr_pymts_df['payment_amount'] > 0].copy()

        # Filter tr_letters_df for pymt_method 'Check'
        tr_df = tr_letters_df.loc[tr_letters_df['pymt_method'] == 'Check'].copy()

        if len(tr_df.index) > 0:
            # Create list of Carriers on checks_df
            tr_carriers = []
            tr_carriers = tr_df.carrier.unique()

            # Loop through carriers list
            for tr_carrier in tr_carriers:

                # Variable Defaults
                tr_sql_where_cal = ''
                tr_letters = True
                tr_s_letters = True
                tr_cals = []

                # check for payment_type_id for TR letters by carrier
                tr_letters_df = tr_df.loc[(tr_df['pymt_type_id'] == 1) & (tr_df['carrier'] == tr_carrier)].copy()

                # TR Letter - create WHERE statement
                if len(tr_letters_df.index) > 0:
                    tr_letters = True
                else:
                    # No Letters
                    tr_letters = False

                # check for payment_type_id for Supplemental letters
                tr_s_letters_df = tr_df.loc[(tr_df['pymt_type_id'] == 2) & (tr_df['carrier'] == tr_carrier)].copy()

                # Supplemental Letter
                if len(tr_s_letters_df.index) > 0:
                    tr_s_letters = True
                else:
                    # No Letters
                    tr_s_letters = False

                # Calculations
                # check for calculations for carrier
                tr_calculations_df = tr_df.loc[tr_df['carrier'] == tr_carrier].copy()

                if len(tr_calculations_df.index) > 0:
                    # Multiple Calculation
                    for index, row in tr_calculations_df.iterrows():
                        tr_cals.append(row['claim_id'])

                # TR - Create letters
                if tr_letters == True:
                    # create df for template
                    tr_template_df = tr_letters_df[letter_cols]

                    # Create TR Letters
                    tr_pdf_template = "letters/pdf_templates/TR_letter_template.pdf"
                    position = 1
                    self.gap_letter(tr_template_df, tr_pdf_template, position)
                
                # Supplemental - Create letters
                if tr_s_letters == True:
                    # create df for template
                    tr_s_template_df = tr_s_letters_df[letter_cols]
                    
                    # Create Supplement Letters
                    tr_s_pdf_template = "letters/pdf_templates/TR_letter_template.pdf"
                    position = 2
                    self.gap_letter(tr_s_template_df, tr_s_pdf_template, position)

                # Calculations - Create SQL query, run query, create calculation sheets
                if tr_letters == True or tr_s_letters == True:

                    # create calculation dataframe
                    temp_cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]                             
                    tr_consolidate_df = pd.DataFrame(columns = temp_cols)

                    for tr_cal in tr_cals:
                        # create sql for calculation
                        tr_sql_query = '''
                                SELECT c.claim_nbr, c.loss_date, l.alt_name, l.contact, b.first, b.last, cc.gap_payable,
                                    cl.incp_date, v.nada_value, CONCAT(v.year, ' ', v.make, ' ', v.model) AS vehicle,
                                    cls.description AS loss_type, cc26.description AS primary_carrier, tr.max_benefit,
                                    cc5.amount AS primary_pymt, cc7.amount AS excess_deductible, tr.term, tr.totalrestart_payable,
                                    cc8.amount AS scr, cc9.amount AS clr, cc10.amount AS cdr, cc11.amount AS oref,
                                    cc18.amount AS salvage, cc19.amount AS prior_dmg, 
                                    cc21.amount AS other1_amt, cc21.description AS other1_description,
                                    cc22.amount AS other2_amt, cc22.description AS other2_description, 
                                    cc35.amount AS other3_amt, cc35.description AS other3_description, ca.description AS carrier,
                                    (cc5.amount + cc7.amount + cc18.amount + cc19.amount + cc.gap_payable + cc8.amount + cc8.amount + 
                                    cc9.amount + cc10.amount + cc11.amount + cc21.amount + cc22.amount + cc35.amount) as subtotal
                                FROM claims c
                                INNER JOIN claim_lender l
                                    USING (claim_id)
                                INNER JOIN claim_borrower b
                                    USING (claim_id)
                                INNER JOIN claim_loan cl
                                    USING (claim_id)
                                INNER JOIN claim_calculations cc
                                    USING (claim_id)
                                INNER JOIN claim_vehicle v
                                    USING (claim_id)
                                INNER JOIN claims_loss_type cls
                                    USING (loss_type_id)
                                INNER JOIN carriers ca
                                    ON c.carrier_id = ca.carrier_id                        
                                INNER JOIN claim_checklist AS cc26
                                    ON c.claim_id = cc26.claim_id
                                    AND cc26.checklist_item_id = 26                  
                                INNER JOIN claim_checklist AS cc5
                                    ON c.claim_id = cc5.claim_id
                                    AND cc5.checklist_item_id = 5
                                INNER JOIN claim_checklist AS cc7
                                    ON c.claim_id = cc7.claim_id
                                    AND cc7.checklist_item_id = 7
                                INNER JOIN claim_checklist AS cc8
                                    ON c.claim_id = cc8.claim_id
                                    AND cc8.checklist_item_id = 8
                                INNER JOIN claim_checklist AS cc9
                                    ON c.claim_id = cc9.claim_id
                                    AND cc9.checklist_item_id = 9
                                INNER JOIN claim_checklist AS cc10
                                    ON c.claim_id = cc10.claim_id
                                    AND cc10.checklist_item_id = 10
                                INNER JOIN claim_checklist AS cc11
                                    ON c.claim_id = cc11.claim_id
                                    AND cc11.checklist_item_id = 11      
                                INNER JOIN claim_checklist AS cc18
                                    ON c.claim_id = cc18.claim_id
                                    AND cc18.checklist_item_id = 18
                                INNER JOIN claim_checklist AS cc19
                                    ON c.claim_id = cc19.claim_id
                                    AND cc19.checklist_item_id = 19     
                                INNER JOIN claim_checklist AS cc21
                                    ON c.claim_id = cc21.claim_id
                                    AND cc21.checklist_item_id = 21
                                INNER JOIN claim_checklist AS cc22
                                    ON c.claim_id = cc22.claim_id
                                    AND cc22.checklist_item_id = 22
                                INNER JOIN claim_checklist AS cc35
                                    ON c.claim_id = cc35.claim_id
                                    AND cc35.checklist_item_id = 35
                                INNER JOIN claim_totalrestart AS tr
                                    ON c.claim_id = tr.claim_id
                                WHERE c.claim_id = {claimID};'''.format(claimID = tr_cal)

                        # save query results as DF
                        tr_temp_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', tr_sql_query, 0, 0))

                        # append query result to c_template_df
                        tr_consolidate_df = tr_consolidate_df.append(tr_temp_df)

                    cal_cols = {0:'claim_nbr', 1:'loss_date', 2:'alt_name', 3:'contact', 4:'first', 5:'last', 6:'gap_payable', 7:'incp_date', 8:'nada_value',
                        9:'vehicle', 10:'loss_type', 11:'primary_carrier', 12:'max_benefit', 13:'primary_pymt', 14:'excess_deductible', 15:'term', 16:'totalrestart_payable', 
                        17:'scr', 18:'clr', 19:'cdr', 20:'oref', 21:'salvage', 22:'prior_dmg', 23:'other1_amt', 24:'other1_description',
                        25:'other2_amt', 26:'other2_description', 27:'other3_amt', 28:'other3_description', 29:'carrier', 30:'subtotal'}

                    tr_template_df = tr_consolidate_df.rename(columns = cal_cols)
                    tr_template_df = tr_template_df.reset_index(drop=True)

                    # Create Calculations
                    tr_pdf_template = "letters/pdf_templates/TR_calculation_template.pdf"
                    position = 3
                    self.tr_calculations(tr_template_df, tr_pdf_template, position)

                # create 
                file_list = self.fileList(self.file_staging_dir, file_ext)

                # Concatenated output file
                outfn = f'S:/claims/letters/TOTALRESTART_{self.now.strftime("%Y-%m-%d")}'

                self.ConCat_pdf(file_list, outfn)

                # Remove files from staging directory
                self.clear_dir(self.file_staging_dir, file_ext)

        else: print('No amount greater then 0.')

        # Update toVGC to 1
        self.update_tovgc_1(tr_df)

        ##### GAP CLAIMS PAID VIA ACH AND $0 CLAIMS #####

        statusText = f'\r\nCreating and emailing GAP Letters and Calculations for $0 and claim paid via ACH...'
        self.updateStatusText(statusText)

        # GAP Claims paid via ACH and $0 Claims
        ach_0_df = gap_pymts_df.loc[(gap_pymts_df['payment_amount'] <= 0) | (gap_pymts_df['pymt_method'].str.contains('ACH'))].copy()

        # # attachment directory
        # attachment_dir = 'S:/claims/letters/attachment/'

        # Remove files from staging directory
        self.clear_dir(self.file_staging_dir, file_ext) 
        self.clear_dir(self.attachment_dir, file_ext)

        if len(ach_0_df.index) > 0:
            letter_cols = ['claim_nbr', 'loss_date', 'alt_name', 'contact', 'address1', 'city', 'state', 'zip', 'first', 'last', 'acct_number', 'payment_amount', 'StateDesc', 'StateCode', 'f_lang' ]
            temp_cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42]

            for index, row in ach_0_df.iterrows():
                # Check for GAP or Supplemental
                if row['pymt_type_id'] == 1:
                    # GAP
                    letter_df = ach_0_df.loc[ach_0_df['claim_id'] == row['claim_id']].copy()       
                    template_df = letter_df[letter_cols]
                    pdf_template = "letters/pdf_templates/GAP_letter_template.pdf"
                    position = 1
                    claim_type = 'GAP Claim'
                else:
                    # Supplemental
                    letter_df = ach_0_df.loc[ach_0_df['claim_id'] == row['claim_id']].copy()       
                    template_df = letter_df[letter_cols]
                    pdf_template = "letters/pdf_templates/Supp_letter_template.pdf"
                    position = 2
                    claim_type = 'Supplemental GAP Claim'

                self.gap_letter(template_df, pdf_template, position)

                # Calculation
                c_consolidate_df = pd.DataFrame(columns = temp_cols)

                c_sql_query = '''
                SELECT c.claim_nbr, c.loss_date, l.alt_name, l.contact, b.first, b.last, cc.gap_payable,
                    cl.incp_date, cl.last_payment, cl.interest_rate, cl.amount AS Amt_Fin, cc.balance_last_pay,
                    cc.nbr_of_days, cc.per_day, ROUND(cc.payoff,2) AS payoff, (cc.ltv * 100) AS ltv, cc.covered_fin_amount,
                    (cc.percent_uncovered * 100) AS percent_uncovered, (cc.ltv_limit * 100) AS ltv_limit, v.nada_value, CONCAT(v.year, ' ', v.make, ' ', v.model) AS vehicle,
                    v.deductible, cls.description AS loss_type, cc26.description AS primary_carrier,
                    cc14.amount AS past_due, cc15.amount AS late_fees, cc16.amount AS skip_pymts,
                    cc17.amount AS skip_fees, cc5.amount AS primary_pymt, cc7.amount AS excess_deductible,
                    cc8.amount AS scr, cc9.amount AS clr, cc10.amount AS cdr, cc11.amount AS oref,
                    cc18.amount AS salvage, cc19.amount AS prior_dmg, cc20.amount AS over_ltv,
                    cc21.amount AS other1_amt, cc21.description AS other1_description,
                    cc22.amount AS other2_amt, cc22.description AS other2_description, ca.description AS carrier,
                    (cc22.amount + cc21.amount + cc20.amount + cc19.amount + cc18.amount + cc11.amount + cc10.amount + cc9.amount + 
                    cc8.amount + cc7.amount + cc5.amount + cc17.amount + cc16.amount + cc15.amount + cc14.amount) as subtotal
                FROM claims c
                INNER JOIN claim_lender l
                    USING (claim_id)
                INNER JOIN claim_borrower b
                    USING (claim_id)
                INNER JOIN claim_loan cl
                    USING (claim_id)
                INNER JOIN claim_calculations cc
                    USING (claim_id)
                INNER JOIN claim_vehicle v
                    USING (claim_id)
                INNER JOIN claims_loss_type cls
                    USING (loss_type_id)
                INNER JOIN carriers ca
                    ON c.carrier_id = ca.carrier_id                        
                INNER JOIN claim_checklist AS cc26
                    ON c.claim_id = cc26.claim_id
                    AND cc26.checklist_item_id = 26    
                INNER JOIN claim_checklist AS cc14
                    ON c.claim_id = cc14.claim_id
                    AND cc14.checklist_item_id = 14    
                INNER JOIN claim_checklist AS cc15
                    ON c.claim_id = cc15.claim_id
                    AND cc15.checklist_item_id = 15    
                INNER JOIN claim_checklist AS cc16
                    ON c.claim_id = cc16.claim_id
                    AND cc16.checklist_item_id = 16       
                INNER JOIN claim_checklist AS cc17
                    ON c.claim_id = cc17.claim_id
                    AND cc17.checklist_item_id = 17       
                INNER JOIN claim_checklist AS cc5
                    ON c.claim_id = cc5.claim_id
                    AND cc5.checklist_item_id = 5
                INNER JOIN claim_checklist AS cc7
                    ON c.claim_id = cc7.claim_id
                    AND cc7.checklist_item_id = 7
                INNER JOIN claim_checklist AS cc8
                    ON c.claim_id = cc8.claim_id
                    AND cc8.checklist_item_id = 8
                INNER JOIN claim_checklist AS cc9
                    ON c.claim_id = cc9.claim_id
                    AND cc9.checklist_item_id = 9
                INNER JOIN claim_checklist AS cc10
                    ON c.claim_id = cc10.claim_id
                    AND cc10.checklist_item_id = 10
                INNER JOIN claim_checklist AS cc11
                    ON c.claim_id = cc11.claim_id
                    AND cc11.checklist_item_id = 11      
                INNER JOIN claim_checklist AS cc18
                    ON c.claim_id = cc18.claim_id
                    AND cc18.checklist_item_id = 18
                INNER JOIN claim_checklist AS cc19
                    ON c.claim_id = cc19.claim_id
                    AND cc19.checklist_item_id = 19
                INNER JOIN claim_checklist AS cc20
                    ON c.claim_id = cc20.claim_id
                    AND cc20.checklist_item_id = 20     
                INNER JOIN claim_checklist AS cc21
                    ON c.claim_id = cc21.claim_id
                    AND cc21.checklist_item_id = 21
                INNER JOIN claim_checklist AS cc22
                    ON c.claim_id = cc22.claim_id
                    AND cc22.checklist_item_id = 22  
                WHERE c.claim_id = {claimID};'''.format(claimID = row['claim_id'])

                # # save query results as DF
                c_temp_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', c_sql_query, 0, 0))

                # append query result to c_template_df
                c_consolidate_df = c_consolidate_df.append(c_temp_df)

                cal_cols = {0:'claim_nbr', 1:'loss_date', 2:'alt_name', 3:'contact', 4:'first', 5:'last', 6:'gap_payable', 7:'incp_date', 8:'last_payment', 9:'interest_rate', 10:'Amt_Fin',
                    11:'balance_last_pay', 12:'nbr_of_days', 13:'per_day', 14:'payoff', 15:'ltv', 16:'covered_fin_amount', 17:'percent_uncovered', 18:'ltv_limit', 19:'nada_value',
                    20:'vehicle', 21:'deductible', 22:'loss_type', 23:'primary_carrier', 24:'past_due', 25:'late_fees', 26:'skip_pymts', 27:'skip_fees', 28:'primary_pymt',
                    29:'excess_deductible', 30:'scr', 31:'clr', 32:'cdr', 33:'oref', 34:'salvage', 35:'prior_dmg', 36:'over_ltv', 37:'other1_amt', 38:'other1_description',
                    39:'other2_amt', 40:'other2_description', 41:'carrier', 42:'subtotal'}

                c_template_df = c_consolidate_df.rename(columns = cal_cols)
                c_template_df = c_template_df.reset_index(drop=True)

                # Create Calculations
                c_pdf_template = "letters/pdf_templates/GAP_calculation_template.pdf"
                position = 3
                self.calculations(c_template_df, c_pdf_template, position)

                # create a list of file to concatenate
                file_list = self.fileList(self.file_staging_dir, file_ext)

                # Concatenated output file
                outfn = self.attachment_dir + row['claim_nbr'] + '-' + row['last']
                
                self.ConCat_pdf(file_list, outfn)

                # send Email
                attach_name = row['claim_nbr'] + '-' + row['last'] + ".pdf"
                missingEmail = 0
                toEmail = []

                if row['email'] == '' and row['email2'] == '':
                    toEmail.append('claims@visualgap.com')
                    missingEmail = 1
                elif row['email'] == '':
                    toEmail.append(row['email2'])
                elif row['email2'] == '':
                    toEmail.append(row['email'])
                else:
                    toEmail.append(row['email'])
                    toEmail.append(row['email2'])

                if missingEmail == 0:
                    msg_html = """
                    <html>
                    <body>
                        <p>Hello {}:<br>
                        <br>
                        We completed the {} for {}'s loan.  Attached is claim letter and calculation. <br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                    """.format(row['contact'], claim_type, (row['first'] + ' ' + row['last']))
                else:
                    msg_html = """
                    <html>
                    <body>
                        <p>There is not an email address on claim {}:<br>
                        <br>
                        Please fax or mail the attached to {}. <br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                    """.format(row['claim_nbr'], row['alt_name'])

                self.send_email(toEmail, claim_type, msg_html, self.attachment_dir, attach_name)

                # Remove files from staging directory
                self.clear_dir(self.attachment_dir, file_ext)
                self.clear_dir(self.file_staging_dir, file_ext)     

        else: print('No ACH and/or $0 claims.')

        # Update toVGC to 1
        self.update_tovgc_1(ach_0_df)

        statusText = f'\r\nCreating and emailing GAP Plus Letters paid via ACH...'
        self.updateStatusText(statusText)

        # Plus Claims paid via ACH and $0 Claims
        ach_plus_df = pymts_df.loc[(pymts_df['payment_category_id'] == 2) & ((pymts_df['payment_amount'] <= 0) | (pymts_df['pymt_method'].str.contains('ACH')))].copy()

        # Remove files from staging directory
        self.clear_dir(self.file_staging_dir, file_ext) 
        self.clear_dir(self.attachment_dir, file_ext)

        letter_cols = ['claim_nbr', 'loss_date', 'alt_name', 'contact', 'address1', 'city', 'state', 'zip', 'first', 'last', 'acct_number', 'payment_amount', 'StateDesc', 'StateCode', 'f_lang' ]

        if len(ach_plus_df.index) > 0:

            for index, row in ach_plus_df.iterrows():

                # create df for template
                letter_df = ach_plus_df.loc[ach_plus_df['claim_id'] == row['claim_id']].copy()
                template_df = letter_df[letter_cols]
                pdf_template = "letters/pdf_templates/PLUS_letter_template.pdf"
                position = 1
                claim_type = "GAP Plus Claim"

                # Create Plus Letter
                self.gap_letter(template_df, pdf_template, position)

                s_dir = f"{self.file_staging_dir}{row['claim_nbr']}-{position}.pdf"
                a_dir = f"{self.attachment_dir}{row['claim_nbr']}-{row['last']}.pdf"
                # Move and rename file to Attachment directory   
                shutil.move(s_dir, a_dir)

                # send Email
                attach_name = f"{row['claim_nbr']}-{row['last']}.pdf"
                missingEmail = 0
                toEmail = []

                if row['email'] == '' and row['email2'] == '':
                    toEmail.append('claims@visualgap.com')
                    missingEmail = 1
                elif row['email'] == '':
                    toEmail.append(row['email2'])
                elif row['email2'] == '':
                    toEmail.append(row['email'])
                else:
                    toEmail.append(row['email'])
                    toEmail.append(row['email2'])

                if missingEmail == 0:
                    msg_html = """
                    <html>
                    <body>
                        <p>Hello {}:<br>
                        <br>
                        We completed the {} for {}'s loan.  Attached is claim letter and calculation. <br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                    """.format(row['contact'], claim_type, (row['first'] + ' ' + row['last']))
                else:
                    msg_html = """
                    <html>
                    <body>
                        <p>There is not an email address on claim {}:<br>
                        <br>
                        Please fax or mail the attached to {}. <br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                    """.format(row['claim_nbr'], row['alt_name'])

                self.send_email(toEmail, claim_type, msg_html, self.attachment_dir, attach_name)

                # Remove files from staging directory
                self.clear_dir(self.attachment_dir, file_ext)

        else: print('No ACH Plus.')

        # Update toVGC to 1
        self.update_tovgc_1(ach_plus_df)

        statusText = f'\r\nCreating and emailing TotalRestart Letters and Calculations for $0 and paid via ACH...'
        self.updateStatusText(statusText)

        # TR Claims paid via ACH and $0 Claims
        tr_ach_0_df = pymts_df.loc[(pymts_df['payment_category_id'] == 3) & ((pymts_df['payment_amount'] <= 0) | (pymts_df['pymt_method'].str.contains('ACH')))].copy()

        # Remove files from staging directory
        self.clear_dir(self.file_staging_dir, file_ext) 
        self.clear_dir(self.attachment_dir, file_ext)

        if len(tr_ach_0_df.index) > 0:
            letter_cols = ['claim_nbr', 'loss_date', 'alt_name', 'contact', 'address1', 'city', 'state', 'zip', 'first', 'last', 'acct_number', 'payment_amount', 'StateDesc', 'StateCode', 'f_lang' ]
            temp_cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42]

            for index, row in tr_ach_0_df.iterrows():
                # Check for GAP or Supplemental
                if row['pymt_type_id'] == 1:
                    # GAP
                    letter_df = tr_ach_0_df.loc[tr_ach_0_df['claim_id'] == row['claim_id']].copy()       
                    template_df = letter_df[letter_cols]
                    pdf_template = "letters/pdf_templates/TR_letter_template.pdf"
                    position = 1
                    claim_type = 'TotalRestart Claim'
                else:
                    # Supplemental
                    letter_df = ach_0_df.loc[ach_0_df['claim_id'] == row['claim_id']].copy()       
                    template_df = letter_df[letter_cols]
                    pdf_template = "letters/pdf_templates/TR_letter_template.pdf"
                    position = 2
                    claim_type = 'Supplemental TotalRestart Claim'

                self.gap_letter(template_df, pdf_template, position)

                # Calculation
                c_consolidate_df = pd.DataFrame(columns = temp_cols)

                c_sql_query = '''
                        SELECT c.claim_nbr, c.loss_date, l.alt_name, l.contact, b.first, b.last, cc.gap_payable,
                            cl.incp_date, v.nada_value, CONCAT(v.year, ' ', v.make, ' ', v.model) AS vehicle,
                            cls.description AS loss_type, cc26.description AS primary_carrier, tr.max_benefit,
                            cc5.amount AS primary_pymt, cc7.amount AS excess_deductible, tr.term, tr.totalrestart_payable,
                            cc8.amount AS scr, cc9.amount AS clr, cc10.amount AS cdr, cc11.amount AS oref,
                            cc18.amount AS salvage, cc19.amount AS prior_dmg, 
                            cc21.amount AS other1_amt, cc21.description AS other1_description,
                            cc22.amount AS other2_amt, cc22.description AS other2_description, 
                            cc35.amount AS other3_amt, cc35.description AS other3_description, ca.description AS carrier,
                            (cc5.amount + cc7.amount + cc18.amount + cc19.amount + cc.gap_payable + cc8.amount + cc8.amount + 
                            cc9.amount + cc10.amount + cc11.amount + cc21.amount + cc22.amount + cc35.amount) as subtotal
                        FROM claims c
                        INNER JOIN claim_lender l
                            USING (claim_id)
                        INNER JOIN claim_borrower b
                            USING (claim_id)
                        INNER JOIN claim_loan cl
                            USING (claim_id)
                        INNER JOIN claim_calculations cc
                            USING (claim_id)
                        INNER JOIN claim_vehicle v
                            USING (claim_id)
                        INNER JOIN claims_loss_type cls
                            USING (loss_type_id)
                        INNER JOIN carriers ca
                            ON c.carrier_id = ca.carrier_id                        
                        INNER JOIN claim_checklist AS cc26
                            ON c.claim_id = cc26.claim_id
                            AND cc26.checklist_item_id = 26                  
                        INNER JOIN claim_checklist AS cc5
                            ON c.claim_id = cc5.claim_id
                            AND cc5.checklist_item_id = 5
                        INNER JOIN claim_checklist AS cc7
                            ON c.claim_id = cc7.claim_id
                            AND cc7.checklist_item_id = 7
                        INNER JOIN claim_checklist AS cc8
                            ON c.claim_id = cc8.claim_id
                            AND cc8.checklist_item_id = 8
                        INNER JOIN claim_checklist AS cc9
                            ON c.claim_id = cc9.claim_id
                            AND cc9.checklist_item_id = 9
                        INNER JOIN claim_checklist AS cc10
                            ON c.claim_id = cc10.claim_id
                            AND cc10.checklist_item_id = 10
                        INNER JOIN claim_checklist AS cc11
                            ON c.claim_id = cc11.claim_id
                            AND cc11.checklist_item_id = 11      
                        INNER JOIN claim_checklist AS cc18
                            ON c.claim_id = cc18.claim_id
                            AND cc18.checklist_item_id = 18
                        INNER JOIN claim_checklist AS cc19
                            ON c.claim_id = cc19.claim_id
                            AND cc19.checklist_item_id = 19     
                        INNER JOIN claim_checklist AS cc21
                            ON c.claim_id = cc21.claim_id
                            AND cc21.checklist_item_id = 21
                        INNER JOIN claim_checklist AS cc22
                            ON c.claim_id = cc22.claim_id
                            AND cc22.checklist_item_id = 22
                        INNER JOIN claim_checklist AS cc35
                            ON c.claim_id = cc35.claim_id
                            AND cc35.checklist_item_id = 35
                        INNER JOIN claim_totalrestart AS tr
                            ON c.claim_id = tr.claim_id
                        WHERE c.claim_id = {claimID};'''.format(claimID = row['claim_id'])

                # # save query results as DF
                c_temp_df = pd.DataFrame(self.mysql_q(vgc_u, vgc_pw, vgc_host, 'visualgap_claims', c_sql_query, 0, 0))

                # append query result to c_template_df
                c_consolidate_df = c_consolidate_df.append(c_temp_df)

                cal_cols = {0:'claim_nbr', 1:'loss_date', 2:'alt_name', 3:'contact', 4:'first', 5:'last', 6:'gap_payable', 7:'incp_date', 8:'nada_value',
                    9:'vehicle', 10:'loss_type', 11:'primary_carrier', 12:'max_benefit', 13:'primary_pymt', 14:'excess_deductible', 15:'term', 16:'totalrestart_payable', 
                    17:'scr', 18:'clr', 19:'cdr', 20:'oref', 21:'salvage', 22:'prior_dmg', 23:'other1_amt', 24:'other1_description',
                    25:'other2_amt', 26:'other2_description', 27:'other3_amt', 28:'other3_description', 29:'carrier', 30:'subtotal'}

                c_template_df = c_consolidate_df.rename(columns = cal_cols)
                c_template_df = c_template_df.reset_index(drop=True)

                # Create Calculations
                c_pdf_template = "letters/pdf_templates/TR_calculation_template.pdf"
                position = 3
                self.tr_calculations(c_template_df, c_pdf_template, position)

                # create a list of file to concatenate
                file_list = self.fileList(self.file_staging_dir, file_ext)

                # Concatenated output file
                outfn = self.attachment_dir + row['claim_nbr'] + '-' + row['last']
                
                self.ConCat_pdf(file_list, outfn)

                # send Email
                attach_name = row['claim_nbr'] + '-' + row['last'] + ".pdf"
                missingEmail = 0
                toEmail = []

                if row['email'] == '' and row['email2'] == '':
                    toEmail.append('claims@visualgap.com')
                    missingEmail = 1
                elif row['email'] == '':
                    toEmail.append(row['email2'])
                elif row['email2'] == '':
                    toEmail.append(row['email'])
                else:
                    toEmail.append(row['email'])
                    toEmail.append(row['email2'])

                if missingEmail == 0:
                    msg_html = """
                    <html>
                    <body>
                        <p>Hello {}:<br>
                        <br>
                        We completed the {} for {}'s loan.  Attached is claim letter and calculation. <br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                    """.format(row['contact'], claim_type, (row['first'] + ' ' + row['last']))
                else:
                    msg_html = """
                    <html>
                    <body>
                        <p>There is not an email address on claim {}:<br>
                        <br>
                        Please fax or mail the attached to {}. <br>
                        <br>
                        Thank you, <br>
                        Claims Department <br>
                        <br>
                        <b>Frost Financial Services, Inc. | VisualGAP <br>
                        Claims Department <br>
                        Phone: 888-753-7678 Option 3</b>
                        </p>
                    </body>
                    </html>
                    """.format(row['claim_nbr'], row['alt_name'])

                self.send_email(toEmail, claim_type, msg_html, self.attachment_dir, attach_name)

                # Remove files from staging directory
                self.clear_dir(self.attachment_dir, file_ext)
                self.clear_dir(self.file_staging_dir, file_ext)

        else: print('No ACH and/or $0 claims.')

        # Update toVGC to 1
        self.update_tovgc_1(tr_ach_0_df)
        statusText = f'\r\nUpdating payment status...'
        self.updateStatusText(statusText)

        statusText = f'\r\nComplete!'
        self.updateStatusText(statusText)

    def qbCustomers(self):
        # Clear Output
        self.clearStatusText()

        # Connect to Quickbooks
        statusText = f'Connecting to Quickbooks...'
        self.updateStatusText(statusText)

        try:
            sessionManager = wc.Dispatch("QBXMLRP2.RequestProcessor")    
            sessionManager.OpenConnection('', 'Claim Payments')
            ticket = sessionManager.BeginSession("", 2)
        except Exception as e:
            print('''
            Make sure to print checks and process ACH in Quickbooks prior to starting this process.
            ERROR: {}'''.format(e))

            statusText = f'\r\nThere was an ERROR connecting to QuickBooks.\r\n***Make sure QuickBooks desktop is open and you are logged into the correct Company file.***\r\nProcess Cancelled.'
            self.updateStatusText(statusText)

            return

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
        try:
            sessionManager = wc.Dispatch("QBXMLRP2.RequestProcessor")    
            sessionManager.OpenConnection('', 'Claim Payments')
            ticket = sessionManager.BeginSession("", 2)
        except Exception as e:
            print('''
            Make sure to print checks and process ACH in Quickbooks prior to starting this process.
            ERROR: {}'''.format(e))

            statusText = f'\r\nThere was an ERROR connecting to QuickBooks.\r\n***Make sure QuickBooks desktop is open and you are logged into the correct Company file.***\r\nProcess Cancelled.'
            self.updateStatusText(statusText)

            return

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
