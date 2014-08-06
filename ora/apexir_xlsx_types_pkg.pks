CREATE OR REPLACE PACKAGE "APEXIR_XLSX_TYPES_PKG" 
  AUTHID CURRENT_USER
AS 
/**
* Package defines all relevant types for APEXIR XLSX Download
* @headcom
*/

  /** 
  * Record with data about a single highlight definition.
  * @param bg_color        Background color to be applied
  * @param font_color      Color to applied to text
  * @param highlight_name  User defined name for highlighting
  * @param highlight_sql   SQL clause to check if highlight should be applied
  * @param col_num         Defines the sql column number to check for highlight being active
  * @param affected_column Name of SQL column affected by highlight
  */
  TYPE t_apex_ir_highlight IS RECORD
    ( bg_color apex_application_page_ir_cond.highlight_row_color%TYPE
    , font_color apex_application_page_ir_cond.highlight_row_font_color%TYPE
    , highlight_name apex_application_page_ir_cond.condition_name%TYPE
    , highlight_sql apex_application_page_ir_cond.condition_sql%TYPE
    , col_num NUMBER
    , affected_column VARCHAR2(30)
    )
  ;
  
  /** Table holding enabled highlights by name */
  TYPE t_apex_ir_highlights IS TABLE OF t_apex_ir_highlight INDEX BY VARCHAR2(30);
  
  /** Table holds active highlights by fetch count */
  TYPE t_apex_ir_active_hl IS TABLE OF t_apex_ir_highlight INDEX BY PLS_INTEGER;

  /** Table holds order number of aggregate enabled column indexed by column name for a single aggregate type */
  TYPE t_apex_ir_aggregate IS TABLE OF PLS_INTEGER INDEX BY VARCHAR2(32767);

  /*
  * Record of all aggregates using tables indexed by column name.
  * Used as temporary storage before attaching aggregates to columns.
  * Allows easy check if column is aggregate enabled by using exists.
  * @param sum_cols             Table of all Columns with "Sum"-Aggregation
  * @param avg_cols             Table of all Columns with "Average"-Aggregation
  * @param max_cols             Table of all Columns with "Max"-Aggregation
  * @param min_cols             Table of all Columns with "Min"-Aggregation
  * @param median_cols          Table of all Columns with "Median"-Aggregation
  * @param count_cols           Table of all Columns with "Count"-Aggregation
  * @param count_distinct_cols  Table of all Columns with "Count Distinct"-Aggregation
  */
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

  /** 
  * Record holds aggregate data attached to an APEX IR column.
  * @param col_num    SQL column number for the aggregate value
  * @param last_value Last seen value of aggregation 
  */
  TYPE t_apex_ir_col_aggregate IS RECORD
    ( col_num PLS_INTEGER
    , last_value NUMBER
    )
  ;
  
  /** Table holds defined aggregates per column indexed by aggregate name. */
  TYPE t_apex_ir_col_aggregates IS TABLE OF t_apex_ir_col_aggregate INDEX BY VARCHAR2(30);

  /**
  * Holds the relevant information about a column in the interactive report.
  * @param report_label      Column heading as defined on intercative report
  * @param is_visible        Defines if column is printed in file.
  * @param is_break_col      Set to TRUE if column is used for control break.
  * @param aggregates        All defined aggregates for the column.
  * @param highlight_conds   All highlights defined on column.
  * @param format_mask       Format mask as defined in IR column definition.
  * @param sql_col_num       SQL column number to retrieve value.
  * @param display_column    Column number in file to show value.
  * @param group_by_function Aggregate Function SQL for Group By View
  */
  TYPE t_apex_ir_col IS RECORD
    ( report_label apex_application_page_ir_col.report_label%TYPE
    , is_visible BOOLEAN
    , is_break_col BOOLEAN := FALSE
    , aggregates t_apex_ir_col_aggregates
    , highlight_conds t_apex_ir_highlights
    , format_mask apex_application_page_ir_col.format_mask%TYPE
    , sql_col_num NUMBER
    , display_column PLS_INTEGER
    , group_by_function apex_application_page_ir_grpby.function_01%TYPE
    )
  ;
  
  /** Table holds all APEX IR columns based on the SQL column alias. */
  TYPE t_apex_ir_cols IS TABLE OF t_apex_ir_col INDEX BY VARCHAR2(30);
  
  /** Stores all used aggregate types for the interactive report by aggregate name */
  TYPE t_apex_ir_active_aggregates IS TABLE OF BOOLEAN INDEX BY VARCHAR2(30);

  /**
  * Holds general information regarding the interactive report.
  * @param application_id             Application the IR belongs to.
  * @param page_id                    Page ID the IR belongs to.
  * @param region_id                  Region ID of the IR region.
  * @param session_id                 Session ID of the request.
  * @param request                    Value of the request variable.
  * @param base_report_id             Report ID for the request.
  * @param report_title               Derived report title.
  * @param report_definition          Report definition as retrieved by calling APEX_IR.GET_REPORT function.
  * @param final_sql                  Final statement used to get the data.
  * @param break_def_column           SQL column number of column with break definition.
  * @param aggregates_offset          SQL column offset to calculate aggregate column numbers.
  * @param active_aggregates          All active aggregate types on IR.
  * @param aggregate_type_disp_column Display column for aggregate types.
  * @param view_mode                  Current view mode of report
  */
  TYPE t_apex_ir_info IS RECORD
    ( application_id NUMBER
    , page_id NUMBER
    , region_id NUMBER
    , session_id NUMBER
    , request VARCHAR2(4000)
    , base_report_id NUMBER
    , report_title VARCHAR2(4000)
    , report_definition apex_ir.t_report
    , final_sql VARCHAR2(32767)
    , break_def_column PLS_INTEGER
    , aggregates_offset PLS_INTEGER
    , active_aggregates t_apex_ir_active_aggregates
    , aggregate_type_disp_column PLS_INTEGER := 1
    , view_mode VARCHAR2(255)
    , group_by_cols VARCHAR2(4000)
    , group_by_sort VARCHAR2(4000)
    , group_by_funcs VARCHAR2(4000)
    )
  ;

  /**
  * Holds the selected options and calculated settings for the XLSX generation.
  * @param show_title            Show header line with report title
  * @param show_filters          Show header lines with filter settings
  * @param show_column_headers   Show column headers before data
  * @param process_highlights    Format data according to highlights
  * @param show_highlights       Show header lines with highlight settings, not useful if set without above
  * @param show_aggregates       Process aggregates and show on total lines
  * @param display_column_count  Holds count of displayed columns, used for merged cells in header section
  * @param sheet                 Holds the worksheet reference
  * @param default_font default  Font for printed values
  * @param default_border_color  Default color if a border is shown
  * @param allow_wrap_text       Switch to allow/disallow word wrap
  * @param original_line_break   The line break character as used in the statement
  * @param replace_line_break    Line break to be used in XLSX file
  * @param default_date_format   Default date format, taken from v$nls_parameters
  * @param append_date_file_name Append current date to file name or not
  * @param requested_view_mode   Interactive Report view mode to use
  */
  TYPE t_xlsx_options IS RECORD
    ( show_title BOOLEAN
    , show_filters BOOLEAN
    , show_column_headers BOOLEAN
    , process_highlights BOOLEAN
    , show_highlights BOOLEAN
    , show_aggregates BOOLEAN
    , display_column_count NUMBER
    , sheet PLS_INTEGER
    , default_font VARCHAR2(100)
    , default_border_color VARCHAR2(100)
    , allow_wrap_text BOOLEAN
    , original_line_break VARCHAR2(10)
    , replace_line_break VARCHAR2(10)
    , default_date_format VARCHAR2(100)
    , append_date_file_name BOOLEAN
    , requested_view_mode VARCHAR2(8)
    )
  ;

  /**
  * Holds column information for all selected columns.
  * @param col_name      Column alias as used in SQL statement.
  * @param col_data_type Data type of the column.
  * @param col_type      Internal column type, e.g. DISPLAY, ROW_HIGHLIGHT, BREAK_DEF...
  * @param is_displayed  Defines if column value should be fetched and put into file.
  */
  TYPE t_sql_col_info IS RECORD
    ( col_name VARCHAR2(32767)
    , col_data_type VARCHAR2(30)
    , col_type VARCHAR2(30)
    , is_displayed BOOLEAN := FALSE
    );
  
  /** Table of all SQL columns by position in SQL statement */
  TYPE t_sql_col_infos IS TABLE OF t_sql_col_info INDEX BY PLS_INTEGER;

  /** Table holding control break offset by SQL row count */
  TYPE t_break_rows IS TABLE OF NUMBER INDEX BY PLS_INTEGER;

  /**
  * Holds general information about the opened cursor, including runtime data.
  * @param cursor_id      ID of opened cursor.
  * @param column_count   Amount of selected columns.
  * @param date_tab       Temporary storage for values of DATE columns.
  * @param num_tab        Temporary storage for values of NUMBER columns.
  * @param vc_tab         Temporary storage for values of VARCHAR columns.
  * @param clob_tab       Temporary storage for values of CLOB columns.
  * @param break_rows     Definition when a control break accours including respective offset for further rows.
  * @param prev_break_val Last seen value of break definition column, needed if bulk size matches break.
  */
  TYPE t_cursor_info IS RECORD
    ( cursor_id PLS_INTEGER
    , column_count PLS_INTEGER
    , date_tab dbms_sql.date_table
    , num_tab dbms_sql.number_table
    , vc_tab dbms_sql.varchar2_table
    , clob_tab dbms_sql.clob_table
    , break_rows t_break_rows
    , prev_break_val VARCHAR2(32767)
    );

  /**
  * Record used to return file and corresponding data.
  * @param file_name The generated file name.
  * @param file_content The generated XLSX file.
  * @param mime_type MIME type to be used for download.
  * @param file_size Size of generated XLSX file.
  */
  TYPE t_returnvalue IS RECORD
    ( file_name VARCHAR2(255)
    , file_content BLOB
    , mime_type VARCHAR2(255)
    , file_size NUMBER
    , error_encountered BOOLEAN := FALSE
    );
    
END APEXIR_XLSX_TYPES_PKG;
/