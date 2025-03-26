# Python Victron Modbus Excel Scan

This is a python script, that uses the Victron Modbus register sheet (excel file) to scan all available registers on a CCGX device and show their actual values based on their number format

Additional python dependencies:

- pandas
- requests
- pymodbus
- numpy
- openpyxl

Pre Operational dependencies:

- Modify the 'myVictronModbusRegisters.txt' file based on your available modbus device IDs.<br>
  On the GX device go to Settings --> Services --> Modbus TCP --> Available Services.<br>
  From there, enter the "com.XXX.XXX" entries with the id values in the txt file line by line (formatting "com.xxx.xxx,id").<br>
  Example: "com.victronenergy.grid,41".<br>
- Victron Excel Sheet (normally downloaded by the script at the beginning)
- User acces to the Victron GX device via modbus tcp in the same subnet, yol'll need the IP address

Then start scanning :)<br>
Tested with the latest Victron CCGX Modbus register list (Rev 50) and python 3.13.2
