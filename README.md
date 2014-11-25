APEXIR_XLSX
===========

Download APEX Interactive Reports as XLSX files.  
With this package you can download any Interactive report using the standard Excel file format.  
Main benefits are that all data types are preserved and stored in a way that Excel recognizes them properly.

SUPPORTED FUNCTIONALITY
-----------------------
*  Filtering and Sorting
*  Control Breaks
*  Computations
*  Aggregations (with a small limitation, see below)
*  Highlights
*  VARCHAR2 columns
*  DATE columns including formatting (supports &APP_DATE_TIME_FORMAT. substitution)
*  NUMBER columns including formatting
*  CLOBs (limited to 32767 characters)
*  "Group By" view mode (for limitations see below)
*  Display column help text as comment on respective column header.

CURRENT LIMITATIONS
-------------------
1. CLOB columns are supported but converted to VARCHAR2(32767) before inserted into the spreadsheet.  
   Support for full size CLOBs is planned for a future release.
2. TIMESTAMP columns are treated like DATE columns.  
   Full support for TIMESTAMP is planned for a future release.
2. Aggregates defined on the first column of any report are ignored.  
   The cell is needed to display the aggregation type.  
   Currently there are no immediate plans to lift that restriction.
3. For the "Group By" view report to work all used columns also need to be present in the "Standard" report view.  
   The APEX engine only exposes the SQL query of the standard view currently.  
4. The download only works for authenticated users.  
   The APEX engine does not give back base_report_id if user isn't authenticated.  
   Research is currently ongoing to lift this restriction.
   
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
   apexir_xlsx_pkg.download( p_ir_region_id => $APEXIR_REGION_ID$);
``` 

###Enable download for all interactive reports in application  
1. Create an application item to hold the requested interactive report id. 
   We'll assume the item is called APEXIR_REGION_ID in the following.
2. Create Application Computation  
   Computation Item: APEXIR_REGION_ID  
   Computation Point: Before Header  
   Computation Type: SQL Query (return single value)  
   Condition Type: Request != Expression 1  
   Expression 1: XLSX  
   Computation (if you have only one IR on the page):

    ```sql
    SELECT region_id
      FROM apex_application_page_ir
     WHERE application_id = :APP_ID
       AND page_id = :APP_PAGE_ID
    ```  

3. Create Application Process  
   Type: PL/SQL anonymous block  
   Process Point: On Load - Before Header  
   Condition Type: Request = Expression 1  
   Expression 1: XLSX  
   Process:

    ```sql
    apexir_xlsx_pkg.download( p_ir_region_id => :APEXIR_REGION_ID);
    ```

4. Create button on page with interactive report.  
   Action: Redirect to Page in this Application  
   Page: The page you are on.  
   Request: XLSX  
