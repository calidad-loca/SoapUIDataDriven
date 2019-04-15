# SoapUIDataDriven
Reads test data and writes it back out to Excel F2C temp conversion

WARNING - This is my first time using GitHub other than the tutorial.
I will put code in different files - hopefully that how this work.

This is a data-driven framework for SoapUI Open Source version (tested on 4.5.0 and 5.5.0) that will read data from an Excel sheet (.xls - not .xlsx), populate the request, send the request, read the response, read the expected result from the Excel sheet and perform an assertion to determine if it PASSED (data returned matches Expected Result).

This will also write results (and rewrite the test data sent) to the Excel sheet.

This uses a free, public webservice that will convert Fahrenheit to Celcius and vice versa.

Unfortunately, this webservice does not return an error msg if an invalid temperature is sent (e.g. Z or $).
