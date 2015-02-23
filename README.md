APEXIR_XLSX
===========

Download APEX Interactive Reports as XLSX files.  
With this package you can download any Interactive report using the standard Excel file format.  
Main benefits are that all data types are preserved and stored in a way that Excel recognizes them properly.

SUPPORTED FUNCTIONALITY
-----------------------
*  Automatically derive Interactive Report Region ID (if only one IR on the page)
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
   This has been identified as a bug by the APEX development Team and should be fixed with APEX 5.0 according to current information.
5. Deriving IR only works with exactly one IR on the page.  
   The package will put a message into the debug log if no or multiple IR regions are found.
   
INSTALLATION
------------
Navigate to folder "setup".
Simply run the script "install_all.sql" if you want everything installed at once.  
If you have the referenced libraries already you can run "install_main.sql" to install the main package.  
Libraries can be installed standalone by using "install_libs.sql".

You can also manually create the packages by running the separate package specifications and bodies in your favourite IDE.

HOW TO USE
----------
###Enable download with default options

1. Create the interactive report region. (skip if you already have one)
2. Create button or similar element to reload page setting request to "XLSX".
3. Create page process (sample)  
   Type: PL/SQL anonymous block  
   Process Point: On Load - Before Header 
   Condition Type: Request = Expression 1  
   Expression 1: XLSX  
   Process:  
```sql
   apexir_xlsx_pkg.download();
``` 

###Setting options (refer to package header for more)

1. Disable help text on column headers  
    ```sql
    apexir_xlsx_pkg.download( p_col_hdr_help => FALSE );
    ```  

2. Do not append date to file name  
    ```sql
    apexir_xlsx_pkg.download( p_append_date => FALSE );
    ```
