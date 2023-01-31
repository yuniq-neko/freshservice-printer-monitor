> **Notice:**
> This script has been long abandoned and is now Archived. I no longer manage bulk amounts of printers and have no need to use this script anymore. If I ever do, I'm sure I'll probably massively re-write this script before bring it back to life. 😉
# freshservice-printer-monitor
A fairly simple script thats designed to be run every 5 minutes or so from Task Scheduler on a Windows PC
Simply does a SNMP snoop on all printers listed in "EPM_data.csv", updates the ink values in the csv, then sends freshservice tickets for each one that's below 1% (Configurable of course!) 😁
