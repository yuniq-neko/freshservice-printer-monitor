# freshservice-printer-monitor
A fairly simple script thats designed to be run every 5 minutes or so from Task Scheduler on a Windows PC
Simply does a SNMP snoop on all printers listed in "EPM_data.csv", updates the ink values in the csv, then sends freshservice tickets for each one that's below 1% (Configurable of course!) 😁
