CREATE OR REPLACE PACKAGE "APEXIR_XLSX_PKG" 
  AUTHID CURRENT_USER
AS
/**
* Package prepares and APEX Interactive Report to be downloaded as an XLSX file.<br />
* Takes all user and global settings and hands them over to XLSX Builder Package by Anton Scheffer.<br />
* Retrieves generated file, amends file name, mime type and file size.
* @headcom
*/

  /** 
  * Retrieve an interactive report as an XLSX file.
  * @param p_ir_region_id        The region ID of the interactive report to convert to XLSX
  * @param p_app_id              Application ID the interactive report belongs to. (derived from APEX context if not manually set)
  * @param p_ir_page_id          ID of page on which the interactive report resides. (derived from APEX context if not set manually)
  * @param p_ir_session_id       APEX session from which to take the session variables. (derived from APEX context if not set manually)
  * @param p_ir_request          Request associated with call. (derived from APEX context if not set manually)
  * @param p_column_headers      Determines if column headers should be rendered. Default: TRUE
  * @param p_aggregates          Determines if aggregates should be rendered. Default: TRUE
  * @param p_process_highlights  Determines if highlights should be considered to color rows and cells. Default: TRUE
  * @param p_show_report_title   Determines if a report title should be rendered as a headline. Default: TRUE
  * @param p_show_filters        Determines if active filters should be rendered as headlines. Default: TRUE
  * @param p_show_highlights     Determines if highlight definitions shoul be rendered as headlines. Default: TRUE
  * @param p_original_line_break Set to the line break used for normal display of the interactive report. Default: &lt;br /&gt;
  * @param p_replace_line_break  Sets the line break used in the XLSX file, replaces original line break set above. Default: \r\n
  * @param p_append_date         Determines if the current date (Format: YYYYMMDD) should be appended to the generated file name. Default: TRUE
  * @return Record Type with file name, generated file, mime type, file size
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
