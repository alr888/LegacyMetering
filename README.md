"# LegacyMetering" 

Setup Guide:

1. Uncompress the zip file to C:\LMG (ensure there is no subdirectory)
2. Open  LMG_Database and view the "LMG Main Form"
3. I have prepared a PR010 file (REQ000000000001.txt) to use. These are the MERI ICPs you have provided you can use the existing file content or edit the ICP list.  
If you want to upload this file click the button "Upload PR010"
4. This is expected to generate a PR030 (EDA file named: EDA000000000001.txt).  
To download it click "Download EDA File"
5. Once the EDA file has been downloaded, you need to click the "External Data" tab and run the "Saved Imports: Import-EDA"
6. After the EDA data has been successfully imported run the query QRY_INSERT_H_RECS.
8. Then run the query QRY_INSERT_I_RECS.
9. Then run the query QRY_INSERT_M_RECS.
10. Again, click on the "External Data" tab and run the "Saved Exports: Export-MEP_LMG":Inline image 5
11. The previous step will generate the MEP update file MEPLMGL00000001.txt which you can then upload to the Registry by clicking the button "Upload MEP File"

As a prototype exception handling and filenames have initially been hardcoded but will be tidied up once we push for full dev't.
Also, there is minor bug in the MEP file generation wherein the header is stuck at the bottom.  Still haven't had the time to fix it but for now if you could edit the MEP file (moving the header back on top) prior to uploading to the Registry it should work.
