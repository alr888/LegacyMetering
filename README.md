"# LegacyMetering" 

Setup Guide:

1. Uncompress the zip file to C:\LMG (ensure there is no subdirectory)
2. Open  LMG_Database and view the "LMG Main Form"
3. I have prepared a PR010 file (REQ000000000001.txt) to use. 
These are the MERI ICPs you have provided you can use the existing file content or edit the ICP list.  
If you want to upload this file click the button "Upload PR010"
4. This is expected to generate a PR030 (EDA file named: EDA000000000001.txt).  
To download it click "Download EDA File"
5. Once the EDA file has been downloaded, you need to click the "External Data" tab and run the "Saved Imports: Import-EDA"
6. Run the query QRY_INSERT_H_RECS.
8. Then run the query QRY_INSERT_I_RECS.
9. Then run the query QRY_INSERT_M_RECS.
10. Then run the query QRY_INSERT_R_RECS.
11. To update the Certificate Flag and Dates, you need to run QRY_UPDATE_CERT_I_RECS.
12. To add/edit/delete any meter details, open the form MEP_UP_METER_DETAILS.
13. To generate the header record, you need to run QRY_INSERT_HEADER.
14. Again, click on the "External Data" tab and run the "Saved Exports: Export-MEP_LMG"
15. Step 10 will generate the MEP update file MEPLMGL00000001.txt which you can then upload to the Registry by clicking the button "Upload MEP File"

Note: The I records containing certificate flag (maps to MEP.DET_3) and dates has  the field DET_5 as the start date 