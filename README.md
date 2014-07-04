APEXIR_XLSX
===========

Download APEX Interactive Reports as XLSX files.

INSTALLATION
------------
Navigate to folder "setup".
Simply run the script "install_all.sql" if you want everything installed at once.
If you have the referenced libraries already you can run "install_main.sql" to install the main package.
Libraries can be installed standalone by using "install_libs.sql".

You can also manually create the packages by running the separate package specifications and bodies in your favourite IDE.

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
   Process (replace $APEXIR_REGION_ID$ with number from step above):
```sql
   DECLARE
     l_xlsx apexir_xlsx_types_pkg.t_returnvalue;
   BEGIN
     l_xlsx := apexir_xlsx_pkg.apexir2sheet( p_ir_region_id => $APEXIR_REGION_ID$);
     OWA_UTIL.mime_header (l_xlsx.mime_type, FALSE);
     HTP.p ('Content-length: ' || l_xlsx.file_size);
     HTP.p ('Content-Disposition: attachment; filename="' || l_xlsx.file_name || '"');
     OWA_UTIL.http_header_close;
     WPG_DOCLOAD.download_file (l_xlsx.file_content);
   END;
``` 

###Enable download for all interactive reports in application  
1. Create an application item to hold the requestes interactive report id. 
   We'll assume the item is called APEXIR_REGION_ID in the following.
2. Create button on page with interactive report. 
   Set Action to "Defined by Dynamic Action" 
   Put following in button attributes:
```
   onclick="javascript:redirect('f?p=&APP_ID.:&APP_PAGE_ID.:&SESSION.:XLSX:&DEBUG.::APEXIR_REGION_ID:' + apex.jQuery('#apexir_REGION_ID').val().substr(1));"
```
   This will reload the page setting request to "XLSX" and APEXIR_REGION_ID application item to the respective region id.
3. Create Application Process 
   Type: PL/SQL anonymous block  
   Process Point: On Load - Before Header 
   Condition Type: Request = Expression 1  
   Expression 1: XLSX  
   Process: 
```sql
   DECLARE
     l_xlsx apexir_xlsx_types_pkg.t_returnvalue;
   BEGIN
     l_xlsx := apexir_xlsx_pkg.apexir2sheet( p_ir_region_id => :APEXIR_REGION_ID);
     OWA_UTIL.mime_header (l_xlsx.mime_type, FALSE);
     HTP.p ('Content-length: ' || l_xlsx.file_size);
     HTP.p ('Content-Disposition: attachment; filename="' || l_xlsx.file_name || '"');
     OWA_UTIL.http_header_close;
     WPG_DOCLOAD.download_file (l_xlsx.file_content);
   END;
``` 
