APEXIR_XLSX
===========

Download APEX Interactive Reports as XLSX files.

INSTALLATION
------------
Simply run the script install_all.sql from folder setup.
You can also manually create the packages by running all files in folder ora in your favorite IDE.

HOW TO USE
----------
###Enable download for a single interactive report

1. Create the interactive report region. (skip if you already have one)
2. Run the page and inspect the source code using your browser.
3. Mark down the value of the hidden item with id "apexir_REGION_ID" removing the "R" at the beginning.  
   The numeric value you now have is the only mandatory input to the main function.
4. Create button or similar element to reload page setting request to "XLSX".
5. Create page process (sample)  
   Type: PL/SQL anonymous block  
   Process Point: On Load - Before Header 
   Condition Type: Request = Expression 1  
   Expression 1: XLSX  
   Process (replace apexir_REGION_ID with number from step above):
```sql
   DECLARE
     v_length NUMBER;
     l_file BLOB;
   BEGIN
     l_file := apexir_xlsx_pkg.apexir2sheet( p_ir_region_id => apexir_REGION_ID);
     v_length := dbms_lob.getlength(l_file);
     OWA_UTIL.mime_header ('application/octet', FALSE);
     HTP.p ('Content-length: ' || v_length);
     HTP.p ( 'Content-Disposition: attachment; filename="apexir_download.xlsx"');
     OWA_UTIL.http_header_close;
     WPG_DOCLOAD.download_file (l_file);
   END;
``` 

###Enable download for all interactive reports in application  
1. Create hidden item on page zero to hold interactive report id. 
   Do not set the item to protected. 
   We'll assume the item is called P0_APEXIR_REGION_ID in the following.
2. Create dynamic action with two true actions.  
   First one sets value of hidden item with javascript and second pins value into session.  
   Javascript Code to set value of item P0_APEXIR_REGION_ID:
   ```javascript
   $('#apexir_REGION_ID').val().substring(1);
   ``` 

   For PL/SQL just put "NULL;" submitting item P0_APEXIR_REGION_ID 
3. Create button on page with interactive report. 
   Set Action to "Redirect to URL" 
   Set Request to XLSX
4. Create Application Process 
   Type: PL/SQL anonymous block  
   Process Point: On Load - Before Header 
   Condition Type: Request = Expression 1  
   Expression 1: XLSX  
   Process (replace apexir_REGION_ID with number from step above): 
```sql
   DECLARE
     v_length NUMBER;
     l_file BLOB;
   BEGIN
     l_file := apexir_xlsx_pkg.apexir2sheet( p_ir_region_id => :P0_APEXIR_REGION_ID);
     v_length := dbms_lob.getlength(l_file);
     OWA_UTIL.mime_header ('application/octet', FALSE);
     HTP.p ('Content-length: ' || v_length);
     HTP.p ( 'Content-Disposition: attachment; filename="apexir_download.xlsx"');
     OWA_UTIL.http_header_close;
     WPG_DOCLOAD.download_file (l_file);
   END;
``` 
