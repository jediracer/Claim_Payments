# Claim_Payments

## Overview
This is project is still in development.  

### Purpose
- Streamline the payment of claims between a proprietary claims system, VisualGAP Claims, and QuickBooks

## Resources
- Software
	- VS Code 1.63, Python 3.8.12 32bit, QuickBooks 2021, QbXml 14.1
- Data Sources
	- MySQL, MS SQL Server, QuickBooks Company Data File stored on local network
	
## Summary
- Collect data from multiple sources via SQL for claims ready to be paid
- Dynamically write QbXml queries to insert payment data into QuickBooks
- Create claim letter and calculation sheets to be mailed with checks
- Email claim letter and calculation sheets for claim paid via ACH
- Update claim system with payment issuance date and check number