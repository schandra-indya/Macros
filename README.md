# Macros
The files uploaded here are macro files for automating excel based reports no matter how complex they are.

P&L Statement Generator:

This project was developed in VBA to automate the generation of automatic P&L report in Excel depicting the health of all the projects running in CTS. The manual effort of developing this report in excel was usually 3 months with high rate of errors. After developing this report, the Company was able to pull out the report in just under 2 minutes for all the sites PAN India.


Tripsheet Preparation Tool:

During the initial phase of implementing the web based transport solution at our Client side, there was a need of creating the tripsheets for employee transport to and fro from Company. So I created this macro, wherein raw excel data was pasted into the third sheet of this macro and then in the tool sheet, you just need to select which tripsheet type (Pick or Drop) is to be prepared. Then the name of the site is to be selected, depending upon which site's data was pasted in the Data Dump sheet. Then how many tripsheets are to be printed in one page needs to be selected. Last but not the least, depending upon Client's requirements, you need to select if the phone number is to be printed in the tripsheet or not and then press the "Generate Report" button. And all the tripsheets are prepared in Tripsheet sheet within 2 minutes. Manual preparation of these tripsheets would take more than one day.


Get Distance Tool:

This tool works in two ways; the first is if you know the latitude and longitude of both origin and destination places. Just place these coordinates in the respective columns. Eg. put Latitude and Longitude of origin in Latitude 1 and Longitude 1 and for destination, put Latitude and Longitude in Latitude 2 and Longitude 2. Then press "Get Distance" button. If the Latitude and Longitude are not known, then put Address of origin in Column A and Address of destination in Column C. Then press "Get Distance" to get the distance between Origin and Destination.


Get Geo Codes Tool:

This tool generates the geo codes (Latitude and Longitude) based on the addresses provided. Just put the addresses in Column A and press "Get Geo Codes" to get the geo codes.


Reverse Geo Coding Tool:

This tool generates the addresses based on the geo codes provided. Just put the Latitude in Column A and Longitude in Column B and press "Get Address" to get the human readable address.
