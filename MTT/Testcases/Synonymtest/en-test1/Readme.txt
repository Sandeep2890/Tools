Triggering get request and verify the responses with expected(Actual) responses

Input: configuration file path
Output: Excel file 

1. Create text file with utterances. 

2. Create text file with expected responses for the utterances. 

3. Create Excel file to record the results of the comparison of responses. 
 
4. Update the configuration file, refer:"configure file"

5. Run the command pip install requests

6. Run the commond pip install openpyxl

7. Run the command 'python synonymtest.py'  'path of configuration file' 
    Ex: python synonymtest.py  "D:\configure.txt"

8. result will be updated in Excel file and text file path provided