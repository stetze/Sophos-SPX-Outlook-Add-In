# Sophos-SPX-Outlook-Add-In



Small Outlook VSTO add-in that adds an Encrypt button to the Outlook ribbon.

The button sets a custom MAPI header used for Sophos Secure Message (SPX) encryption workflows.



**Function**

When the Encrypt button is activated, the add-in writes the following custom header to the mail item:



SPX\_HEADER = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/x-sophos-spx-encrypt"



This header can be evaluated by downstream mail security solutions (e.g. Sophos Email) to trigger encryption.



**Notes**



Implemented as Outlook VSTO Add-In

Designed for enterprise environments using Sophos SPX



**License**

See LICENSE.txt



\## Outlook Ribbon

!\[image](Assets/Activated.png)

!\[image](Assets/Default.png)

