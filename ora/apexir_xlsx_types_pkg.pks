CREATE OR REPLACE PACKAGE "APEXIR_XLSX_TYPES_PKG" 
  AUTHID CURRENT_USER
AS 

  -- Holds relevant information about enabled highlights
  TYPE t_apex_ir_highlight IS RECORD
    ( bg_color apex_application_page_ir_cond.highlight_row_color%TYPE
    , font_color apex_application_page_ir_cond.highlight_row_font_color%TYPE
    , highlight_name apex_application_page_ir_cond.condition_name%TYPE
    , highlight_sql apex_application_page_ir_cond.condition_sql%TYPE
    , col_num NUMBER -- defines which SQL column to check
    , affected_column VARCHAR2(30)
    )
  ;
  
  TYPE t_apex_ir_highlights IS TABLE OF t_apex_ir_highlight INDEX BY VARCHAR2(30);
  TYPE t_apex_ir_active_hl IS TABLE OF t_apex_ir_highlight INDEX BY PLS_INTEGER;

  TYPE t_apex_ir_aggregate IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(32767);

  TYPE t_apex_ir_aggregates IS RECORD
    ( sum_cols t_apex_ir_aggregate
    , avg_cols t_apex_ir_aggregate
    , max_cols t_apex_ir_aggregate
    , min_cols t_apex_ir_aggregate
    , median_cols t_apex_ir_aggregate
    , count_cols t_apex_ir_aggregate
    , count_distinct_cols t_apex_ir_aggregate    
    )
  ;

  TYPE t_apex_ir_col_aggregate IS RECORD
    ( col_num PLS_INTEGER
    , last_value NUMBER
    )
  ;
  TYPE t_apex_ir_col_aggregates IS TABLE OF t_apex_ir_col_aggregate INDEX BY VARCHAR2(30);

  -- Holds the relevant information about a column in the interactive report.
  TYPE t_apex_ir_col IS RECORD
    ( report_label apex_application_page_ir_col.report_label%TYPE -- Column heading
    , is_visible BOOLEAN -- Defines if column should be visible in XLSX file
    , is_break_col BOOLEAN -- Defines if column is used to determine control break
    , aggregates t_apex_ir_col_aggregates -- Aggregates defined on the column
    , highlight_conds t_apex_ir_highlights -- Highlights defined on the column
    , format_mask apex_application_page_ir_col.format_mask%TYPE -- Format mask
    , sql_col_num NUMBER -- defines which SQL column to check
    , display_column PLS_INTEGER -- Column number in XLSX file
    )
  ;
  
  TYPE t_apex_ir_cols IS TABLE OF t_apex_ir_col INDEX BY VARCHAR2(30);
  
  TYPE t_apex_ir_active_aggregates IS TABLE OF BOOLEAN INDEX BY VARCHAR2(30);

  -- Holds general information regarding the interactive report.
  TYPE t_apex_ir_info IS RECORD
    ( application_id NUMBER -- Application ID IR belongs to
    , page_id NUMBER -- Page ID IR belongs to
    , region_id NUMBER -- Region ID of IR Region
    , session_id NUMBER -- Session ID for Request
    , request VARCHAR2(4000) -- Request value
    , base_report_id NUMBER -- Report ID for Request
    , report_title VARCHAR2(4000) -- Derived Report Title
    , report_definition apex_ir.t_report -- Collected using APEX function APEX_IR.GET_REPORT
    , final_sql VARCHAR2(32767) -- Final SQL statement used to get all data
    , break_def_column PLS_INTEGER -- sql column number of break definition
    , aggregates_offset PLS_INTEGER -- sql column offset when calculating aggregate column numbers
    , active_aggregates t_apex_ir_active_aggregates -- which types of aggregates are active ( count gives row offset )
    , aggregate_type_disp_column PLS_INTEGER := 1 -- defaults to 1, meaning aggregates active but no break columns
    )
  ;

  -- Holds the selected options and calculated settings for the XLSX generation.
  TYPE t_xlsx_options IS RECORD
    ( show_title BOOLEAN -- Show header line with report title
    , show_filters BOOLEAN -- show header lines with filter settings
    , show_column_headers BOOLEAN -- show column headers before data
    , process_highlights BOOLEAN -- format data according to highlights
    , show_highlights BOOLEAN -- show header lines with highlight settings, not useful if set without above
    , show_aggregates BOOLEAN -- process aggregates and show on total lines
    , display_column_count NUMBER -- holds the count of displayed columns, used for merged cells in header section
    , sheet PLS_INTEGER -- holds the worksheet reference
    , default_font VARCHAR2(100) -- default font for printed values
    , default_border_color VARCHAR2(100) -- default color if a border is shown
    , allow_wrap_text BOOLEAN -- Switch to allow/disallow word wrap
    , original_line_break VARCHAR2(10) -- The line break character as used in the statement
    , replace_line_break VARCHAR2(10) -- Line break to be used in XLSX file
    , default_date_format VARCHAR2(100) -- Default date format, taken from v$nls_parameters
    , append_date_file_name BOOLEAN -- Append current date to file name or not
    )
  ;

  -- Holds column information for all selected columns.  
  TYPE t_sql_col_info IS RECORD
    ( col_name VARCHAR2(32767) -- Column alias in SQL statement
    , col_data_type VARCHAR2(30) -- Data Type of the column
    , col_type VARCHAR2(30) -- Internal column type, e.g. DISPLAY, ROW_HIGHLIGHT, BREAK_DEF
    , is_displayed BOOLEAN := FALSE -- assume no display, we loop through all and flag displayed then
    );
  
  TYPE t_sql_col_infos IS TABLE OF t_sql_col_info INDEX BY PLS_INTEGER;

  TYPE t_break_rows IS TABLE OF NUMBER INDEX BY PLS_INTEGER; --zero based break rows

  -- Holds general information about the opened cursor, including runtime data.  
  TYPE t_cursor_info IS RECORD
    ( cursor_id PLS_INTEGER -- ID of opened cursor
    , column_count PLS_INTEGER -- Count of selected columns
    , current_row PLS_INTEGER -- Current SQL row
    , date_tab dbms_sql.date_table
    , num_tab dbms_sql.number_table
    , vc_tab dbms_sql.varchar2_table
    , clob_tab dbms_sql.clob_table
    , break_rows t_break_rows -- Definition of when a break occures including respective offset
    , prev_break_val VARCHAR2(32767) -- Previous value of break column, needed if bulk size matches break
    );

  -- Holds the data to return.    
  TYPE t_returnvalue IS RECORD
    ( file_name VARCHAR2(255) -- Generated filename
    , file_content BLOB -- The XLSX file
    , mime_type VARCHAR2(255) -- mime type to be used for download
    , file_size NUMBER -- size of the XLSX file
    );
    
END APEXIR_XLSX_TYPES_PKG;
/