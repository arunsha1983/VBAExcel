How about DNS Forward and Reverse lookup as well as Ping!

Press Alt-F11 in Excel to get to the VBA screen.
Right click on the Project View
Click Add Module
Add the following snippet.
Use:
Hide   Copy Code
GetHostname("4.2.2.1") in any Excel cell.<br />
or<br />
Use: <code>GetIpAddress("www.google.com") in any Excel cell.<br />
or
Use:
Hide   Copy Code
Ping("4.2.2.1") in any Excel cell.<br />


Note: If you have lengthy lists of IP addresses yoou plan on looking up keep in mind that the processs for looking up things against DNS is SLOW in comparison to a normal formula within Excel. Be wise with how you utilize this function.

In cases with duplicate addresses such as a ‘ip flow-cache’ output or ‘ip accounting’ output, you should probably create one lookup table in Excel on a separate tab with a deduplicated list of the IP Addresses with you will attempt loookup in DNS against. Then just use VLOOKUP(IP,HostLookupTable,Col,FALSE) to update the main page.

Once looked up I always copy and paste-as-text so excel doesn’t constantly lookup the list.

Credits: Many Thanks to AlonHircsh and Arkham79 for leaving little gem on experts-echange. It has been modified to include Arkham’s suggestion of including conversion of longIpAddress to stringIpAddress in the function.
