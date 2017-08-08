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
  * @param p_ir_region_id        The region ID of the interactive report to convert to XLSX. (derived from APEX context if not set manually)
  * @param p_app_id              Application ID the interactive report belongs to. (derived from APEX context if not set manually)
  * @param p_ir_page_id          ID of page on which the interactive report resides. (derived from APEX context if not set manually)
  * @param p_ir_session_id       APEX session from which to take the session variables. (derived from APEX context if not set manually)
  * @param p_ir_request          Request associated with call. (derived from APEX context if not set manually)
  * @param p_ir_view_mode        Sets the interactive report view mode to use.
  *                              Leave NULL to use APEX context.
  * @param p_column_headers      Determines if column headers should be rendered. Default: TRUE
  * @param p_col_hdr_help        Determines if the help text is added as a comment to column headers. Default: TRUE
  * @param p_aggregates          Determines if aggregates should be rendered. Default: TRUE
  * @param p_process_highlights  Determines if highlights should be considered to color rows and cells. Default: TRUE
  * @param p_show_report_title   Determines if a report title should be rendered as a headline. Default: TRUE
  * @param p_show_filters        Determines if active filters should be rendered as headlines. Default: TRUE
  * @param p_include_page_items  Determines if used page items should be rendered as headlines. Default: FALSE
  * @param p_show_highlights     Determines if highlight definitions should be rendered as headlines. Default: TRUE
  * @param p_original_line_break Set to the line break used for normal display of the interactive report. Default: &lt;br /&gt;
  * @param p_replace_line_break  Sets the line break used in the XLSX file, replaces original line break set above. Default: \r\n
  * @param p_filter_replacement  Sets the value to be used when replacing original line break in filter display row. Default is one blank
  * @param p_append_date         Determines if the current date (Format: YYYYMMDD) should be appended to the generated file name. Default: TRUE
  * @param p_append_date_fmt     The date format to be used when appending current date to filename. Default: YYYYMMDD
  * @return Record Type with file name, generated file, mime type, file size
  */
  FUNCTION apexir2sheet
    ( p_ir_region_id NUMBER := NULL
    , p_app_id NUMBER := NV('APP_ID')
    , p_ir_page_id NUMBER := NV('APP_PAGE_ID')
    , p_ir_session_id NUMBER := NV('SESSION')
    , p_ir_request VARCHAR2 := V('REQUEST')
    , p_ir_view_mode VARCHAR2 := NULL
    , p_column_headers BOOLEAN := TRUE
    , p_col_hdr_help BOOLEAN := TRUE
    , p_freeze_col_hdr BOOLEAN := FALSE
    , p_aggregates IN BOOLEAN := TRUE
    , p_process_highlights IN BOOLEAN := TRUE
    , p_show_report_title IN BOOLEAN := TRUE
    , p_show_filters IN BOOLEAN := TRUE
    , p_include_page_items IN BOOLEAN := FALSE
    , p_show_highlights IN BOOLEAN := TRUE
    , p_original_line_break IN VARCHAR2 := '<br />'
    , p_replace_line_break IN VARCHAR2 := chr(13) || chr(10)
    , p_filter_replacement IN VARCHAR2 := ' '
    , p_append_date IN BOOLEAN := TRUE
    , p_append_date_fmt IN VARCHAR2 := 'YYYYMMDD'
    )
  RETURN apexir_xlsx_types_pkg.t_returnvalue;

  /**
  * Download Interactive Report as XLSX file.
  * This is a wrapper for APEXIR2SHEET which immediately presents the file for download.
  */

  PROCEDURE download
    ( p_ir_region_id NUMBER := NULL
    , p_app_id NUMBER := NV('APP_ID')
    , p_ir_page_id NUMBER := NV('APP_PAGE_ID')
    , p_ir_session_id NUMBER := NV('SESSION')
    , p_ir_request VARCHAR2 := V('REQUEST')
    , p_ir_view_mode VARCHAR2 := NULL
    , p_column_headers BOOLEAN := TRUE
    , p_col_hdr_help BOOLEAN := TRUE
    , p_freeze_col_hdr BOOLEAN := FALSE
    , p_aggregates IN BOOLEAN := TRUE
    , p_process_highlights IN BOOLEAN := TRUE
    , p_show_report_title IN BOOLEAN := TRUE
    , p_show_filters IN BOOLEAN := TRUE
    , p_include_page_items IN BOOLEAN := FALSE
    , p_show_highlights IN BOOLEAN := TRUE
    , p_original_line_break IN VARCHAR2 := '<br />'
    , p_replace_line_break IN VARCHAR2 := chr(13) || chr(10)
    , p_filter_replacement IN VARCHAR2 := ' '
    , p_append_date IN BOOLEAN := TRUE
    , p_append_date_fmt IN VARCHAR2 := 'YYYYMMDD'
    );

  FUNCTION get_version
    RETURN VARCHAR2;

END APEXIR_XLSX_PKG;

/
