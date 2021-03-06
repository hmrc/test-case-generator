
# test-case-generator

This application is designed to be used in conjunction with the MTR Tester spreadsheet that is provided by the business.
The latest document can be found here: https://confluence.tools.tax.service.gov.uk/pages/viewpage.action?pageId=160893275

The purpose of the application is to create three new sheets in the MTR tester workbook. These sheets point to various other
parts of the spreadsheet to pull input, interim and output values that can subsequently be used to generate CSVs that can be used
in the SA Filing application integration tests.

Before executing the application download the latest MTR Tester from the page above. If you are not working on Windows 
DO NOT open the spreadsheet file as it is likely to be corrupted if it is opened using Libre Office or any other non-Windows application.
Make sure that you read this Confluence page before proceeding: https://confluence.tools.tax.service.gov.uk/display/SAF1617/SA-Filing+-+Calculation+Engine+Testing
as it explains the process of preparing the MTR Tester spreadsheet prior to using the test case generator application.

### Usage

The application expects two parameters, and input file which is the MTR tester and an output file name which is where the
altered file will be saved. Currently the application expects the input file to exist in the resources folder and the output file will be written here as well.

To run the application use:
<pre>
 ./gradle run --args='-i [input-file-name] -o [output-file-name]'
</pre>


### License

This code is open source software licensed under the [Apache 2.0 License]("http://www.apache.org/licenses/LICENSE-2.0.html").
