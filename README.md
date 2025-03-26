# Python Victron Modbus Excel Scan

This is a python script, that uses the Victron Modbus register sheet (excel file) to scan all available registers and show their actual values based on their number format

Additional python dependencies:

- pandas
- requests
- pymodbus
- numpy
- openpyxl

Operational dependencies:

- Victron Excel Sheet (normally downloaded by the script at the beginning)
- User acces to the Victron GX device via modbus tcp in the same subnet, IP address

Tested with the latest Victron CCGX Modbus register list (Rev 50) and python 3.13.2
