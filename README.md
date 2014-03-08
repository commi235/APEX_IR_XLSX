APEX_IR_XLSX
============

Download APEX Interactive Reports as XLSX files.

INSTALLATION
============
Simply run the script install_all.sql from folder setup.
You can also manually create the packages by running all files in folder ora in your favorite IDE.

HOW TO USE
==========
Option 1: Enable for a single interactive report
1. Create the interactive report region. (skip if you already have one)
2. Run the page and inspect the source code using your browser.
3. Mark down the value of the hidden item with id "apexir_REGION_ID" removing the "R" at the beginning.
   The numeric value you now have is the only mandatory input to the main function.
4. Create page process (sample)
   Type: PL/SQL anonymous block
   Process Point: On Load - Before Header
   Conditions
     Condition Type: Request = Expression 1
     Expression 1: XLSX
   Process (replace apexir_REGION_ID with number from step above):
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
5. Create button or similar element to reload page setting request to "XLSX".

Option 2: Enable for all interactive reports
1. TBD
