CREATE OR REPLACE PACKAGE BODY "MK_APEX_IR_XLSX" 
AS

/* Constants */
  c_bulk_size CONSTANT pls_integer := 200;

/* TYPE Definitions */  
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

  -- Types to hold run-time information
  TYPE t_apex_ir_col IS RECORD
    ( report_label apex_application_page_ir_col.report_label%TYPE
    , is_visible BOOLEAN
    , is_break_col BOOLEAN
    , sum_on_break BOOLEAN
    , avg_on_break BOOLEAN
    , max_on_break BOOLEAN
    , min_on_break BOOLEAN
    , median_on_break BOOLEAN
    , count_on_break BOOLEAN
    , count_distinct_on_break BOOLEAN
    , highlight_conds t_apex_ir_highlights
    , sql_col_num NUMBER -- defines which SQL column to check
    , display_column PLS_INTEGER
    )
  ;
  
  TYPE t_apex_ir_cols IS TABLE OF t_apex_ir_col INDEX BY VARCHAR2(30);

  TYPE t_apex_ir_info IS RECORD
    ( application_id NUMBER -- Application ID IR belongs to
    , page_id NUMBER -- Page ID IR belongs to
    , region_id NUMBER -- Region ID of IR Region
    , session_id NUMBER -- Session ID for Request
    , base_report_id NUMBER -- Report ID for Request
    , report_title VARCHAR2(4000) -- Derived Report Title
    , report_definition apex_ir.t_report -- Collected using APEX function APEX_IR.GET_REPORT
    , final_sql VARCHAR2(32767)
    )
  ;

  TYPE t_xlsx_options IS RECORD
    ( show_title BOOLEAN -- Show header line with report title
    , show_filters BOOLEAN -- show header lines with filter settings
    , show_column_headers BOOLEAN -- show column headers before data
    , process_highlights BOOLEAN -- format data according to highlights
    , show_highlights BOOLEAN -- show header lines with highlight settings, not useful if set without above
    , show_aggregates BOOLEAN -- process aggregates and show on total lines
    , display_column_count NUMBER -- holds the count of displayed columns, used for merged cells in header section
    , sheet PLS_INTEGER -- holds the worksheet reference
    , default_font VARCHAR2(100)
    )
  ;
  
/* Global Variables */
 
  -- runtime data
  g_apex_ir_info t_apex_ir_info;
  g_xlsx_options t_xlsx_options;
  g_col_settings t_apex_ir_cols;
  g_row_highlights t_apex_ir_highlights;
  g_col_highlights t_apex_ir_highlights;
  g_current_row PLS_INTEGER := 1;

/* Support Procedures */

  PROCEDURE get_report_title
  AS
  BEGIN
    SELECT CASE
             WHEN rpt.report_name IS NOT NULL THEN
               ir.region_name || ' - ' || rpt.report_name
             ELSE
               ir.region_name
           END report_title
      INTO g_apex_ir_info.report_title
      FROM apex_application_page_ir ir JOIN apex_application_page_ir_rpt rpt
             ON ir.application_id = rpt.application_id
            AND ir.page_id = rpt.page_id
            AND ir.interactive_report_id = rpt.interactive_report_id
     WHERE ir.application_id = g_apex_ir_info.application_id
       AND ir.page_id = g_apex_ir_info.page_id
       AND rpt.base_report_id = g_apex_ir_info.base_report_id
       AND rpt.session_id = g_apex_ir_info.session_id
    ;
  END get_report_title;

  PROCEDURE get_std_columns
  AS
    col_rec t_apex_ir_col;
  BEGIN
    -- These column names are static defined, used as reference
    FOR rec IN ( SELECT column_alias, report_label, display_text_as
                   FROM APEX_APPLICATION_PAGE_IR_COL
                  WHERE page_id = g_apex_ir_info.page_id
                    AND application_id = g_apex_ir_info.application_id
                    AND region_id = g_apex_ir_info.region_id )
    LOOP
      col_rec.report_label := rec.report_label;
      col_rec.is_visible := rec.display_text_as != 'HIDDEN';      
      g_col_settings(rec.column_alias) := col_rec;
    END LOOP;
  END get_std_columns;


  PROCEDURE get_computations
  AS
    col_rec  t_apex_ir_col;
  BEGIN
    -- computations are run-time data, therefore need base report ID and session
    FOR rec IN (SELECT computation_column_alias,computation_report_label
                  FROM apex_application_page_ir_comp comp JOIN apex_application_page_ir_rpt rpt
                         ON rpt.application_id = comp.application_id
                        AND rpt.page_id = comp.page_id
                        AND rpt.report_id = comp.report_id
                 WHERE rpt.application_id = g_apex_ir_info.application_id
                   AND rpt.page_id = g_apex_ir_info.page_id
                   AND rpt.base_report_id = g_apex_ir_info.base_report_id
                   AND rpt.session_id = g_apex_ir_info.session_id)
    LOOP
      col_rec.report_label := rec.computation_report_label;
      col_rec.is_visible := TRUE;
      g_col_settings(rec.computation_column_alias) := col_rec;
    END LOOP;
  END get_computations;

  PROCEDURE get_aggregates
  AS
    l_avg_cols  apex_application_page_ir_rpt.avg_columns_on_break%TYPE;
    l_break_on  apex_application_page_ir_rpt.break_enabled_on%TYPE;
    l_count_cols  apex_application_page_ir_rpt.count_columns_on_break%TYPE;
    l_count_distinct_cols  apex_application_page_ir_rpt.count_distnt_col_on_break%TYPE;
    l_cur_col  VARCHAR2(30);
    l_max_cols  apex_application_page_ir_rpt.max_columns_on_break%TYPE;
    l_median_cols  apex_application_page_ir_rpt.median_columns_on_break%TYPE;
    l_min_cols  apex_application_page_ir_rpt.min_columns_on_break%TYPE;
    l_sum_cols  apex_application_page_ir_rpt.sum_columns_on_break%TYPE;
  BEGIN
    -- First get run-time settings for aggregate infos
    SELECT break_enabled_on,
           sum_columns_on_break,
           avg_columns_on_break,
           max_columns_on_break,
           min_columns_on_break,
           median_columns_on_break,
           count_columns_on_break,
           count_distnt_col_on_break
      INTO l_break_on,
           l_sum_cols,
           l_avg_cols,
           l_max_cols,
           l_min_cols,
           l_median_cols,
           l_count_cols,
           l_count_distinct_cols           
      FROM apex_application_page_ir_rpt
     WHERE application_id = g_apex_ir_info.application_id
       AND page_id = g_apex_ir_info.page_id
       AND base_report_id = g_apex_ir_info.base_report_id
       AND session_id = g_apex_ir_info.session_id;

    -- Loop through all selected columns and apply settings
    l_cur_col := g_col_settings.FIRST();
    WHILE (l_cur_col IS NOT NULL)
    LOOP
      g_col_settings(l_cur_col).is_break_col := l_break_on IS NOT NULL AND INSTR(l_break_on, l_cur_col) > 0;
      g_col_settings(l_cur_col).sum_on_break := l_sum_cols IS NOT NULL AND INSTR(l_sum_cols, l_cur_col) > 0;
      g_col_settings(l_cur_col).avg_on_break := l_avg_cols IS NOT NULL AND INSTR(l_avg_cols, l_cur_col) > 0;
      g_col_settings(l_cur_col).max_on_break := l_max_cols IS NOT NULL AND INSTR(l_max_cols, l_cur_col) > 0;
      g_col_settings(l_cur_col).min_on_break := l_min_cols IS NOT NULL AND INSTR(l_min_cols, l_cur_col) > 0;
      g_col_settings(l_cur_col).median_on_break := l_median_cols IS NOT NULL AND INSTR(l_median_cols, l_cur_col) > 0;
      g_col_settings(l_cur_col).count_on_break := l_count_cols IS NOT NULL AND INSTR(l_count_cols, l_cur_col) > 0;
      g_col_settings(l_cur_col).count_distinct_on_break := l_count_distinct_cols IS NOT NULL AND INSTR(l_count_distinct_cols, l_cur_col) > 0;
      l_cur_col := g_col_settings.next(l_cur_col);
    END LOOP;
  END get_aggregates;

  PROCEDURE get_highlights
  AS
    col_rec t_apex_ir_highlight;
    hl_num NUMBER := 0;
  BEGIN
    FOR rec IN (SELECT CASE
                         WHEN cond.highlight_row_color IS NOT NULL OR cond.highlight_row_font_color IS NOT NULL
                           THEN NULL
                         ELSE cond.condition_column_name
                       END condition_column_name,
                       REPLACE (cond.condition_sql, '#APXWS_EXPR#', cond.condition_expression) test_sql,
                       cond.condition_name,
                       REPLACE(COALESCE(cond.highlight_row_color, cond.highlight_cell_color), '#') bg_color,
                       REPLACE(COALESCE(cond.highlight_row_font_color, cond.highlight_cell_font_color), '#') font_color
                  FROM apex_application_page_ir_cond cond JOIN apex_application_page_ir_rpt r
                         ON r.application_id = cond.application_id
                        AND r.page_id = cond.page_id
                        AND r.report_id = cond.report_id
                 WHERE cond.application_id = g_apex_ir_info.application_id
                   AND cond.page_id = g_apex_ir_info.page_id
                   AND cond.condition_type = 'Highlight'
                   AND cond.condition_enabled = 'Yes'
                   AND r.base_report_id = g_apex_ir_info.base_report_id
                   AND r.session_id = g_apex_ir_info.session_id
                   AND ( cond.highlight_row_color IS NOT NULL
                      OR cond.highlight_row_font_color IS NOT NULL
                      OR cond.highlight_cell_color IS NOT NULL
                      OR cond.highlight_cell_font_color IS NOT NULL
                       )
                ORDER BY cond.condition_column_name, cond.highlight_sequence
               )
    LOOP
      hl_num := hl_num + 1;
      col_rec.bg_color := rec.bg_color;
      col_rec.font_color := rec.font_color;
      col_rec.highlight_name := rec.condition_name;
      col_rec.highlight_sql := REPLACE(rec.test_sql, '#APXWS_HL_ID#', 1);
      col_rec.affected_column := rec.condition_column_name;
      IF rec.condition_column_name IS NOT NULL AND g_col_settings.EXISTS(rec.condition_column_name) THEN
        g_col_highlights('HL_' || to_char(hl_num)) := col_rec;
      ELSE
        g_row_highlights('HL_' || to_char(hl_num)) := col_rec;
      END IF;
      g_apex_ir_info.final_sql := g_apex_ir_info.final_sql || ', ' || col_rec.highlight_sql || ' AS HL_' || to_char(hl_num);
    END LOOP;
  END get_highlights;

  PROCEDURE get_settings
  AS
  BEGIN
    get_report_title;
    get_std_columns;
    get_computations;
    IF g_xlsx_options.show_aggregates THEN
      get_aggregates;
    END IF;
    IF g_xlsx_options.process_highlights THEN
      get_highlights;
    END IF;
  END get_settings;

  PROCEDURE print_header
  AS
    l_cur_hl_name VARCHAR2(30);
  BEGIN
    IF g_xlsx_options.show_title THEN
      ax_xlsx_builder.mergecells( p_tl_col => 1
                                , p_tl_row => g_current_row
                                , p_br_col => g_xlsx_options.display_column_count
                                , p_br_row => g_current_row
                                , p_sheet => g_xlsx_options.sheet
                                );
      ax_xlsx_builder.cell( p_col => 1
                          , p_row => g_current_row
                          , p_value => g_apex_ir_info.report_title
                          , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                , p_fontsize => 14
                                                                , p_bold => TRUE
                                                                )
                          , p_alignment => ax_xlsx_builder.get_alignment( p_vertical => 'center'
                                                                          , p_horizontal => 'center'
                                                                          )
                          , p_sheet => g_xlsx_options.sheet
                          );
      g_current_row := g_current_row + 1;
    END IF;
    IF g_xlsx_options.show_filters THEN
      -- TODO Implementation required
      NULL;
/*    
      ax_xlsx_builder.mergecells( p_tl_col => 1
                                , p_tl_row => g_current_row
                                , p_br_col => g_xlsx_options.display_column_count
                                , p_br_row => g_current_row
                                , p_sheet => g_xlsx_options.sheet
                                );
      g_current_row := g_current_row + 1;
*/
    END IF;
    IF g_xlsx_options.show_highlights THEN
      l_cur_hl_name := g_row_highlights.FIRST();
      WHILE (l_cur_hl_name IS NOT NULL) LOOP
        ax_xlsx_builder.mergecells( p_tl_col => 1
                                  , p_tl_row => g_current_row
                                  , p_br_col => g_xlsx_options.display_column_count
                                  , p_br_row => g_current_row
                                  , p_sheet => g_xlsx_options.sheet
                                  );
        ax_xlsx_builder.cell( p_col => 1
                            , p_row => g_current_row
                            , p_value => g_row_highlights(l_cur_hl_name).highlight_name
                            , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                  , p_rgb => g_row_highlights(l_cur_hl_name).font_color
                                                                  )
                            , p_fillId => ax_xlsx_builder.get_fill( p_patternType => 'solid'
                                                                  , p_fgRGB => g_row_highlights(l_cur_hl_name).bg_color
                                                                  )
                            , p_alignment => ax_xlsx_builder.get_alignment( p_vertical => 'center'
                                                                          , p_horizontal => 'center'
                                                                          )
                            , p_sheet => g_xlsx_options.sheet );
        g_current_row := g_current_row + 1;
        l_cur_hl_name := g_row_highlights.next(l_cur_hl_name);
      END LOOP;
      l_cur_hl_name := g_col_highlights.FIRST();
      WHILE (l_cur_hl_name IS NOT NULL) LOOP
        ax_xlsx_builder.mergecells( p_tl_col => 1
                                  , p_tl_row => g_current_row
                                  , p_br_col => g_xlsx_options.display_column_count
                                  , p_br_row => g_current_row
                                  , p_sheet => g_xlsx_options.sheet
                                  );
        ax_xlsx_builder.cell( p_col => 1
                            , p_row => g_current_row
                            , p_value => g_col_highlights(l_cur_hl_name).highlight_name
                            , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                  , p_rgb => g_col_highlights(l_cur_hl_name).font_color
                                                                  )
                            , p_fillId => ax_xlsx_builder.get_fill( p_patternType => 'solid'
                                                                  , p_fgRGB => g_col_highlights(l_cur_hl_name).bg_color
                                                                  )
                            , p_alignment => ax_xlsx_builder.get_alignment( p_vertical => 'center'
                                                                          , p_horizontal => 'center'
                                                                          )
                            , p_sheet => g_xlsx_options.sheet );
        g_current_row := g_current_row + 1;        
        l_cur_hl_name := g_col_highlights.next(l_cur_hl_name);
      END LOOP;
    END IF;
    g_current_row := g_current_row + 1;
  END print_header;
  
  
/* Main Function */

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
  RETURN BLOB
  AS
    l_cursor integer;
    l_col_cnt integer;
    l_desc_tab dbms_sql.desc_tab2;
    l_date_tab dbms_sql.date_table;
    l_num_tab dbms_sql.number_table;
    l_vc_tab dbms_sql.varchar2_table;
    l_hl_tab dbms_sql.number_table;
    t_r INTEGER;
    l_cur_col_name VARCHAR2(4000);
    l_num_val NUMBER;
    l_cur_col_highlight t_apex_ir_highlight;
    l_active_col_highlights t_apex_ir_active_hl;
  BEGIN
    -- IR infos
    g_apex_ir_info.application_id := p_app_id;
    g_apex_ir_info.page_id := p_ir_page_id;
    g_apex_ir_info.session_id := p_ir_session_id;
    g_apex_ir_info.region_id := p_ir_region_id;
    g_apex_ir_info.base_report_id := apex_ir.get_last_viewed_report_id(p_page_id => g_apex_ir_info.page_id, p_region_id => g_apex_ir_info.region_id); -- set manual for test outside APEX Environment
    g_apex_ir_info.report_definition := APEX_IR.GET_REPORT ( p_page_id => g_apex_ir_info.page_id, p_region_id => g_apex_ir_info.region_id);
    
    -- Generation Options
    g_xlsx_options.show_aggregates := p_aggregates;
    g_xlsx_options.process_highlights := p_process_highlights;
    g_xlsx_options.show_title := p_show_report_title;
    g_xlsx_options.show_filters := p_show_filters;
    g_xlsx_options.show_highlights := p_show_highlights;
    g_xlsx_options.show_column_headers := p_column_headers;
    g_xlsx_options.display_column_count := 0; -- shift result set to right if > 0
    g_xlsx_options.default_font := 'Arial';
    g_xlsx_options.sheet := ax_xlsx_builder.new_sheet; -- needed before running any ax_xlsx_builder commands

    -- retrieve IR infos
    get_settings();

    -- Split sql query on first from and inject highlight conditions
    g_apex_ir_info.final_sql := SUBSTR(g_apex_ir_info.report_definition.sql_query, 1, INSTR(UPPER(g_apex_ir_info.report_definition.sql_query), ' FROM')) 
                             || g_apex_ir_info.final_sql
                             || SUBSTR(g_apex_ir_info.report_definition.sql_query, INSTR(UPPER(g_apex_ir_info.report_definition.sql_query), ' FROM'));
    
    l_cursor := dbms_sql.open_cursor;
    dbms_sql.parse( l_cursor, g_apex_ir_info.final_sql, dbms_sql.NATIVE );
    dbms_sql.describe_columns2( l_cursor, l_col_cnt, l_desc_tab );
    
    /* Bind values from IR structure*/
    FOR i IN 1..g_apex_ir_info.report_definition.binds.count LOOP
      dbms_sql.bind_variable( l_cursor, g_apex_ir_info.report_definition.binds(i).name, g_apex_ir_info.report_definition.binds(i).value);
    END LOOP;

    /* Amend column settings and create header row */    
    FOR c IN 1 .. l_col_cnt LOOP
      IF g_col_settings.exists(l_desc_tab(c).col_name) THEN
        IF g_col_settings(l_desc_tab(c).col_name).is_visible THEN -- remove hidden cols
          g_xlsx_options.display_column_count := g_xlsx_options.display_column_count + 1; -- count number of displayed columns
          g_col_settings(l_desc_tab(c).col_name).sql_col_num := c; -- column in SQL
          g_col_settings(l_desc_tab(c).col_name).display_column := g_xlsx_options.display_column_count; -- column in spreadsheet
        END IF;
      ELSIF g_row_highlights.EXISTS(l_desc_tab(c).col_name) THEN
        g_row_highlights(l_desc_tab(c).col_name).col_num := c;
      ELSIF g_col_highlights.EXISTS(l_desc_tab(c).col_name) THEN
        g_col_highlights(l_desc_tab(c).col_name).col_num := c;
        l_cur_col_highlight := g_col_highlights(l_desc_tab(c).col_name);
        g_col_settings(l_cur_col_highlight.affected_column).highlight_conds(l_desc_tab(c).col_name) := l_cur_col_highlight;
      END IF;
      
      CASE
        WHEN l_desc_tab( c ).col_type IN ( 2, 100, 101 ) THEN
          dbms_sql.define_array( l_cursor, c, l_num_tab, c_bulk_size, 1 );
        WHEN l_desc_tab( c ).col_type IN ( 12, 178, 179, 180, 181 , 231 ) THEN
          dbms_sql.define_array( l_cursor, c, l_date_tab, c_bulk_size, 1 );
        WHEN l_desc_tab( c ).col_type IN ( 1, 8, 9, 96, 112 ) THEN
          dbms_sql.define_array( l_cursor, c, l_vc_tab, c_bulk_size, 1 );
        ELSE
          NULL;
      END CASE;
    END LOOP;
    
    IF g_xlsx_options.show_title OR g_xlsx_options.show_filters OR g_xlsx_options.show_highlights THEN
      print_header;
    END IF;
    
    IF g_xlsx_options.show_column_headers THEN
      FOR c IN 1..l_col_cnt LOOP
        IF g_col_settings.EXISTS(l_desc_tab(c).col_name)
       AND g_col_settings(l_desc_tab(c).col_name).is_visible THEN
          ax_xlsx_builder.cell( p_col => g_col_settings(l_desc_tab(c).col_name).display_column
                              , p_row => g_current_row
                              , p_value => g_col_settings(l_desc_tab(c).col_name).report_label
                              , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                    , p_bold => TRUE
                                                                    )
                              , p_sheet => g_xlsx_options.sheet );
        END IF;
      END LOOP;
      g_current_row := g_current_row + 1;
    END IF;

    -- Start looping through the "real" data
    t_r := dbms_sql.execute( l_cursor );
    loop
      t_r := dbms_sql.fetch_rows( l_cursor );
      IF t_r > 0 THEN
        -- first run through row highlights
        l_cur_col_name := g_row_highlights.FIRST();
        WHILE (l_cur_col_name IS NOT NULL) LOOP
          dbms_sql.COLUMN_VALUE( l_cursor, g_row_highlights(l_cur_col_name).col_num, l_num_tab );
          FOR i IN 0 .. t_r - 1 LOOP
            IF (l_num_tab(i + l_num_tab.FIRST()) IS NOT NULL) THEN
              ax_xlsx_builder.set_row( p_row => g_current_row + i
                                     , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                           , p_rgb => g_row_highlights(l_cur_col_name).font_color
                                                                           )
                                     , p_fillId => ax_xlsx_builder.get_fill( p_patternType => 'solid'
                                                                           , p_fgRGB => g_row_highlights(l_cur_col_name).bg_color
                                                                           )
                                     );
            END IF;
          END LOOP;
          l_num_tab.DELETE;
          l_cur_col_name := g_row_highlights.NEXT(l_cur_col_name);
        END LOOP;
        -- run through displayed columns
        FOR c IN 1..l_col_cnt LOOP
          -- new column, clean highlights
          l_active_col_highlights.delete;
          IF g_col_settings.EXISTS(l_desc_tab(c).col_name) THEN
            IF g_col_settings(l_desc_tab(c).col_name).is_visible THEN -- remove hidden cols
              -- check if column has highlights attached
              IF g_col_settings(l_desc_tab(c).col_name).highlight_conds.count() > 0 THEN
                l_cur_col_name := g_col_settings(l_desc_tab(c).col_name).highlight_conds.FIRST;
                WHILE (l_cur_col_name IS NOT NULL) LOOP
                  l_cur_col_highlight := g_col_settings(l_desc_tab(c).col_name).highlight_conds(l_cur_col_name);
                  dbms_sql.COLUMN_VALUE( l_cursor, l_cur_col_highlight.col_num, l_num_tab);
                  FOR i IN 0 .. t_r - 1 LOOP
                    -- highlight condition TRUE
                    IF l_num_tab(i + l_num_tab.FIRST()) IS NOT NULL THEN
                      -- no previous highlight condition matched
                      IF NOT l_active_col_highlights.EXISTS(i) THEN
                        l_active_col_highlights(i) := l_cur_col_highlight;
                      END IF;
                    END IF;
                  END LOOP;
                  -- next record, clean up array and get next index, !don't forget this ever!
                  l_num_tab.delete;
                  l_cur_col_name := g_col_settings(l_desc_tab(c).col_name).highlight_conds.next(l_cur_col_name);
                END LOOP;
              END IF;
              
              -- now create the cells
              CASE
                WHEN l_desc_tab( c ).col_type IN ( 2, 100, 101 ) THEN
                  dbms_sql.COLUMN_VALUE( l_cursor, c, l_num_tab );
                  FOR i IN 0 .. t_r - 1 loop
                    IF l_active_col_highlights.EXISTS(i) THEN
                      ax_xlsx_builder.cell( p_col => g_col_settings(l_desc_tab(c).col_name).display_column
                                          , p_row => g_current_row + i
                                          , p_value => l_num_tab( i + l_num_tab.FIRST() )
                                          , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                                , p_rgb => l_active_col_highlights(i).font_color
                                                                                )
                                          , p_fillId => ax_xlsx_builder.get_fill( p_patternType => 'solid'
                                                                                , p_fgRGB => l_active_col_highlights(i).bg_color
                                                                                )
                                          , p_sheet => g_xlsx_options.sheet );
                    ELSE
                      ax_xlsx_builder.cell( p_col => g_col_settings(l_desc_tab(c).col_name).display_column
                                          , p_row => g_current_row + i
                                          , p_value => l_num_tab( i + l_num_tab.FIRST() )
                                          , p_sheet => g_xlsx_options.sheet );
                    END IF;
                  END loop;
                  l_num_tab.DELETE;
                WHEN l_desc_tab( c ).col_type IN ( 12, 178, 179, 180, 181 , 231 ) THEN
                  dbms_sql.column_value( l_cursor, c, l_date_tab );
                  FOR i IN 0 .. t_r - 1 loop
                    IF l_active_col_highlights.EXISTS(i) THEN
                      ax_xlsx_builder.cell( p_col => g_col_settings(l_desc_tab(c).col_name).display_column
                                          , p_row => g_current_row + i
                                          , p_value => l_date_tab( i + l_date_tab.FIRST() )
                                          , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                                , p_rgb => l_active_col_highlights(i).font_color
                                                                                )
                                          , p_fillId => ax_xlsx_builder.get_fill( p_patternType => 'solid'
                                                                                , p_fgRGB => l_active_col_highlights(i).bg_color
                                                                                )
                                          , p_sheet => g_xlsx_options.sheet );
                    ELSE
                      ax_xlsx_builder.cell( p_col => g_col_settings(l_desc_tab(c).col_name).display_column
                                          , p_row => g_current_row + i
                                          , p_value => l_date_tab( i + l_date_tab.FIRST() )
                                          , p_sheet => g_xlsx_options.sheet );
                    END IF;
                  END loop;
                  l_date_tab.delete;
                WHEN l_desc_tab( c ).col_type IN ( 1, 8, 9, 96, 112 ) THEN
                  dbms_sql.column_value( l_cursor, c, l_vc_tab );
                  FOR i IN 0 .. t_r - 1 loop
                    IF l_active_col_highlights.EXISTS(i) THEN
                      ax_xlsx_builder.cell( p_col => g_col_settings(l_desc_tab(c).col_name).display_column
                                          , p_row => g_current_row + i
                                          , p_value => l_vc_tab( i + l_vc_tab.FIRST() )
                                          , p_alignment => ax_xlsx_builder.get_alignment(p_wrapText => FALSE)
                                          , p_fontId => ax_xlsx_builder.get_font( p_name => g_xlsx_options.default_font
                                                                                , p_rgb => l_active_col_highlights(i).font_color
                                                                                )
                                          , p_fillId => ax_xlsx_builder.get_fill( p_patternType => 'solid'
                                                                                , p_fgRGB => l_active_col_highlights(i).bg_color
                                                                                )
                                          , p_sheet => g_xlsx_options.sheet );
                    ELSE
                      ax_xlsx_builder.cell( p_col => g_col_settings(l_desc_tab(c).col_name).display_column
                                          , p_row => g_current_row + i
                                          , p_value => l_vc_tab( i + l_vc_tab.FIRST() )
                                          , p_alignment => ax_xlsx_builder.get_alignment(p_wrapText => FALSE)
                                          , p_sheet => g_xlsx_options.sheet );
                    END IF;
                  END loop;
                  l_vc_tab.delete;
                else
                  NULL;
              END CASE;
            END IF;
          END IF;
        END LOOP;
      end if;
      exit when t_r != c_bulk_size;
      g_current_row := g_current_row + t_r;
    end loop;
    dbms_sql.close_cursor( l_cursor );
    RETURN ax_xlsx_builder.finish;
  exception
    when others
    then
      if dbms_sql.is_open( l_cursor )
      then
        dbms_sql.close_cursor( l_cursor );
      END IF;
      raise;
      RETURN NULL;
  END query2sheet_apex;

END MK_APEX_IR_XLSX;

/
