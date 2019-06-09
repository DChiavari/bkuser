"bkuser" script
---------------

Author: Danilo Chiavari (www.danilochiavari.com)
Date:   June 9th, 2019

This script was born to facilitate backup policy creations for workstations in Veeam Backup & Replication / Veeam Agent for Windows.

The script does the following:

  -  Parses a list of computers (hostnames or IPs) from a CSV file (`list.csv`, newline-delimited), then for each of them:
  	  -  Retrieves the last logged on user name from the Windows registry
	  -  Sets an environmental variable ("bkuser") in the system context, with the last logged on user name as value

After doing this, it is easy to leverage the `bkuser` environmental variable when creating a backup policy from within Veeam Backup & Replication (or standalone Veeam Agent as well).

For example: in order to back up only the 'Desktop' folders, `C:\Users\%bkuser%\Desktop` can be used.

In order to run the script, you need to have administrative privileges on the machine you are running the script from and on all target workstations.

USAGE: open an elevated Command Prompt (Run as Administrator) and run the script directly: `C:\scripts\bkuser\bkuser.vbs`

Running the script using `cscript` is recommended, as all "logging" (`wscript.echo`) output would be written to console (command prompt window) rather than pop-up boxes.

EXAMPLE: `cscript C:\scripts\bkuser\bkuser.vbs`

![script execution from the backup server](bkuser-script-vbr95.png)
_Script execution from the backup server_

![policy definition using the variable](bkuser-policy-objects-vbr95.png)
_Policy definition using the variable_

![file-level recovery from a backup](bkuser-FLR-win10-01.png)
_File-level recovery from a backup_
