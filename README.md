# Azure2Word.ps1
Azure2Word uses powershell to get Azure information and puts this in a Word document
Works with Word Office365

Connects to Azure powershell to get information
.PARAMETER Customer
[REQUIRED] This is the companyname that is used in the documentation
.PARAMETER ReportPath
[REQUIRED] This is the path where the report is stored
.PARAMETER TenantId
[OPTIONAL] When you have mutiple Tennants (CSP) enter the Tenant GUID
.PARAMETER SubscriptionId
[OPTIONAL] When you have mutiple Subscriptions (CSP) enter the Subscription GUID
.\Azure2Word -Customer Contoso -ReportPath c:\reports
    Runs the script for customer Contoso and create the word file in c:\reports
.\Azure2Word -Customer Contoso -ReportPath c:\reports -TenantId xxxx-xxxx -SubscriptionId yyyy-yyyy
    Runs the script for customer Contoso and create the word file in c:\reports 
    Running the scipt on TenantID xxxx-xxxx and SubscriptionID yyyy-yyy

Author: Xander Angenent (@XaAng70)
Idea: Anders Bengtsson http://contoso.se/blog/?p=4286
Last Modified: 2020/12/4