CREATE OR REPLACE PACKAGE "MK_APEX_IR_XLSX" 
  AUTHID DEFINER
AS 

  /* Feature Set:
      - Highlights for rows and columns
      
    TODO:
      - Header rows for report title and settings
  */

  FUNCTION query2sheet_apex
    ( p_ir_region_id NUMBER
    , p_app_id NUMBER := NV('APP_ID')
    , p_ir_page_id NUMBER := NV('APP_PAGE_ID')
    , p_ir_session_id NUMBER := NV('SESSION')
    , p_column_headers BOOLEAN := TRUE
    , p_aggregates IN BOOLEAN := FALSE
    , p_process_highlights IN BOOLEAN := TRUE
    , p_show_report_title IN BOOLEAN := TRUE
    , p_show_filters IN BOOLEAN := TRUE
    , p_show_highlights IN BOOLEAN := TRUE
    )
  RETURN BLOB;

END MK_APEX_IR_XLSX;

/
