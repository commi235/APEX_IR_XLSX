CREATE OR REPLACE PACKAGE "APEXIR_XLSX_PKG" 
  AUTHID DEFINER
AS 

  /* Feature Set:
      - Highlights for rows and columns
      - Header rows for title, highlights and filters
    
     TODO:
      - Include aggregates for break column and totals
  */

  FUNCTION apexir2sheet
    ( p_ir_region_id NUMBER
    , p_app_id NUMBER := NV('APP_ID')
    , p_ir_page_id NUMBER := NV('APP_PAGE_ID')
    , p_ir_session_id NUMBER := NV('SESSION')
    , p_ir_request VARCHAR2 := V('REQUEST')
    , p_column_headers BOOLEAN := TRUE
    , p_aggregates IN BOOLEAN := TRUE
    , p_process_highlights IN BOOLEAN := TRUE
    , p_show_report_title IN BOOLEAN := TRUE
    , p_show_filters IN BOOLEAN := TRUE
    , p_show_highlights IN BOOLEAN := TRUE
    , p_original_line_break IN VARCHAR2 := '<br />'
    , p_replace_line_break IN VARCHAR2 := chr(13) || chr(10)
    , p_append_date IN BOOLEAN := TRUE
    )
  RETURN apexir_xlsx_types_pkg.t_returnvalue;

END APEXIR_XLSX_PKG;

/
