CREATE OR REPLACE PACKAGE "APEXIR_XLSX_PKG" 
  AUTHID CURRENT_USER
AS 

  /* Feature Set:
      - Highlights for rows and columns
      - Header rows for title, highlights and filters
  */

  FUNCTION apexir2sheet
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

END APEXIR_XLSX_PKG;

/
