CREATE OR REPLACE PACKAGE "XLSX_BUILDER_PKG"
  AUTHID CURRENT_USER
IS
/**********************************************
**
** Author: Anton Scheffer
** Date: 19-02-2011
** Website: http://technology.amis.nl/blog
** See also: http://technology.amis.nl/blog/?p=10995
**
** Changelog:
**   Date: 21-02-2011
**     Added Aligment, horizontal, vertical, wrapText
**   Date: 06-03-2011
**     Added Comments, MergeCells, fixed bug for dependency on NLS-settings
**   Date: 16-03-2011
**     Added bold and italic fonts
**   Date: 22-03-2011
**     Fixed issue with timezone's set to a region(name) instead of a offset
**   Date: 08-04-2011
**     Fixed issue with XML-escaping from text
**   Date: 27-05-2011
**     Added MIT-license
**   Date: 11-08-2011
**     Fixed NLS-issue with column width
**   Date: 29-09-2011
**     Added font color
**   Date: 16-10-2011
**     fixed bug in add_string
**   Date: 26-04-2012
**     Fixed set_autofilter (only one autofilter per sheet, added _xlnm._FilterDatabase)
**     Added list_validation = drop-down 
**   Date: 27-08-2013
**     Added freeze_pane
**   Date: 01-03-2014 (MK)
**     Changed new_sheet to function returning sheet id
**   Date: 22-03-2014 (MK)
**     Added function to convert Oracle Number Format to Excel Format
**   Date: 07-04-2014 (MK)
**     Removed references to UTL_FILE
**     query2sheet is now function returning BLOB
**     changed date handling to be based on 01-01-1900
**   Date: 08-04-2014 (MK)
**     internal function for date to excel serial conversion added
******************************************************************************
******************************************************************************
Copyright (C) 2011, 2012 by Anton Scheffer

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

******************************************************************************
******************************************** */

  TYPE tp_alignment IS RECORD
    ( vertical VARCHAR2(11)
    , horizontal VARCHAR2(16)
    , wrapText BOOLEAN
    )
  ;

  TYPE tp_XF_fmt IS RECORD
    ( numFmtId PLS_INTEGER
    , fontId PLS_INTEGER
    , fillId PLS_INTEGER
    , borderId PLS_INTEGER
    , alignment tp_alignment
    );
  TYPE tp_col_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
  TYPE tp_row_fmts IS TABLE OF tp_XF_fmt INDEX BY PLS_INTEGER;
  TYPE tp_widths IS TABLE OF NUMBER INDEX BY PLS_INTEGER;
  
  TYPE tp_cell IS RECORD
    ( value_id NUMBER
    , style_def VARCHAR2(50)
    );

  TYPE tp_cells IS TABLE OF tp_cell INDEX BY PLS_INTEGER;
  TYPE tp_rows IS TABLE OF tp_cells INDEX BY PLS_INTEGER;

  TYPE tp_autofilter IS RECORD
    ( column_start PLS_INTEGER
    , column_end PLS_INTEGER
    , row_start PLS_INTEGER
    , row_end PLS_INTEGER
    );
  TYPE tp_autofilters IS TABLE OF tp_autofilter INDEX BY PLS_INTEGER;
  TYPE tp_hyperlink IS RECORD
    ( cell VARCHAR2(10)
    , url  VARCHAR2(1000)
    );
  TYPE tp_hyperlinks IS TABLE OF tp_hyperlink INDEX BY PLS_INTEGER;
  subtype tp_author IS VARCHAR2(32767 CHAR);
  type tp_authors is table of PLS_INTEGER index by tp_author;
  authors tp_authors;
  type tp_comment is record
    ( text VARCHAR2(32767 char)
    , author tp_author
    , row PLS_INTEGER
    , column PLS_INTEGER
    , width PLS_INTEGER
    , height PLS_INTEGER
    );
  type tp_comments is table of tp_comment index by PLS_INTEGER;
  type tp_mergecells is table of VARCHAR2(21) index by PLS_INTEGER;
  type tp_validation is record
    ( type VARCHAR2(10)
    , errorstyle VARCHAR2(32)
    , showinputmessage boolean
    , prompt VARCHAR2(32767 char)
    , title VARCHAR2(32767 char)
    , error_title VARCHAR2(32767 char)
    , error_txt VARCHAR2(32767 char)
    , showerrormessage boolean
    , formula1 VARCHAR2(32767 char)
    , formula2 VARCHAR2(32767 char)
    , allowBlank boolean
    , sqref VARCHAR2(32767 char)
    );
  type tp_validations is table of tp_validation index by PLS_INTEGER;
  type tp_sheet is record
    ( rows tp_rows
    , widths tp_widths
    , name VARCHAR2(100)
    , freeze_rows PLS_INTEGER
    , freeze_cols PLS_INTEGER
    , autofilters tp_autofilters
    , hyperlinks tp_hyperlinks
    , col_fmts tp_col_fmts
    , row_fmts tp_row_fmts
    , comments tp_comments
    , mergecells tp_mergecells
    , validations tp_validations
    );
  type tp_sheets is table of tp_sheet index by PLS_INTEGER;
  type tp_numFmt is record
    ( numFmtId PLS_INTEGER
    , formatCode VARCHAR2(100)
    );
  type tp_numFmts is table of tp_numFmt index by PLS_INTEGER;
  type tp_fill is record
    ( patternType VARCHAR2(30)
    , fgRGB VARCHAR2(8)
    );
  type tp_fills is table of tp_fill index by PLS_INTEGER;
  type tp_cellXfs is table of tp_xf_fmt index by PLS_INTEGER;
  type tp_font is record
    ( name VARCHAR2(100)
    , family PLS_INTEGER
    , fontsize number
    , theme PLS_INTEGER
    , RGB VARCHAR2(8)
    , underline boolean
    , italic boolean
    , bold boolean
    );
  type tp_fonts is table of tp_font index by PLS_INTEGER;
  type tp_border is record
    ( top VARCHAR2(17)
    , bottom VARCHAR2(17)
    , left VARCHAR2(17)
    , right VARCHAR2(17)
    );
  type tp_borders is table of tp_border index by PLS_INTEGER;
  type tp_numFmtIndexes is table of PLS_INTEGER index by PLS_INTEGER;
  type tp_strings is table of PLS_INTEGER index by VARCHAR2(32767 char);
  type tp_str_ind is table of VARCHAR2(32767 char) index by PLS_INTEGER;
  type tp_defined_name is record
    ( name VARCHAR2(32767 char)
    , ref VARCHAR2(32767 char)
    , sheet PLS_INTEGER
    );
  type tp_defined_names is table of tp_defined_name index by PLS_INTEGER;
  type tp_book is record
    ( sheets tp_sheets
    , strings tp_strings
    , str_ind tp_str_ind
    , str_cnt PLS_INTEGER := 0
    , fonts tp_fonts
    , fills tp_fills
    , borders tp_borders
    , numFmts tp_numFmts
    , cellXfs tp_cellXfs
    , numFmtIndexes tp_numFmtIndexes
    , defined_names tp_defined_names
    );
--
  FUNCTION get_workbook
    RETURN tp_book;
    
  PROCEDURE clear_workbook;
--
  FUNCTION new_sheet( p_sheetname VARCHAR2 := NULL )
    RETURN PLS_INTEGER;
--
  function OraFmt2Excel( p_format VARCHAR2 := NULL )
  return VARCHAR2;
--

  FUNCTION OraNumFmt2Excel ( p_format VARCHAR2 )
    RETURN VARCHAR2;
    
  function get_numFmt( p_format VARCHAR2 := NULL )
  return PLS_INTEGER;
--
  function get_font
    ( p_name VARCHAR2
    , p_family PLS_INTEGER := 2
    , p_fontsize number := 8
    , p_theme PLS_INTEGER := 1
    , p_underline boolean := false
    , p_italic boolean := false
    , p_bold boolean := FALSE
    , p_rgb VARCHAR2 := NULL -- this is a hex ALPHA Red Green Blue value, but RGB works also
    )
  return PLS_INTEGER;
--
  function get_fill
    ( p_patternType VARCHAR2
    , p_fgRGB VARCHAR2 := NULL -- this is a hex ALPHA Red Green Blue value, but RGB works also
    )
  return PLS_INTEGER;
--
  function get_border
    ( p_top VARCHAR2 := 'thin'
    , p_bottom VARCHAR2 := 'thin'
    , p_left VARCHAR2 := 'thin'
    , p_right VARCHAR2 := 'thin'
    )
/*
none
thin
medium
dashed
dotted
thick
double
hair
mediumDashed
dashDot
mediumDashDot
dashDotDot
mediumDashDotDot
slantDashDot
*/
  return PLS_INTEGER;
--
  function get_alignment
    ( p_vertical VARCHAR2 := NULL
    , p_horizontal VARCHAR2 := NULL
    , p_wrapText boolean := NULL
    )
/* horizontal
center
centerContinuous
distributed
fill
general
justify
left
right
*/
/* vertical
bottom
center
distributed
justify
top
*/
  return tp_alignment;
--
  PROCEDURE cell
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_value number
    , p_numFmtId PLS_INTEGER := NULL
    , p_fontId PLS_INTEGER := NULL
    , p_fillId PLS_INTEGER := NULL
    , p_borderId PLS_INTEGER := NULL
    , p_alignment tp_alignment := NULL
    , p_sheet PLS_INTEGER := NULL
    );
--
  PROCEDURE cell
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_value VARCHAR2
    , p_numFmtId PLS_INTEGER := NULL
    , p_fontId PLS_INTEGER := NULL
    , p_fillId PLS_INTEGER := NULL
    , p_borderId PLS_INTEGER := NULL
    , p_alignment tp_alignment := NULL
    , p_sheet PLS_INTEGER := NULL
    );
--
  PROCEDURE cell
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_value date
    , p_numFmtId PLS_INTEGER := NULL
    , p_fontId PLS_INTEGER := NULL
    , p_fillId PLS_INTEGER := NULL
    , p_borderId PLS_INTEGER := NULL
    , p_alignment tp_alignment := NULL
    , p_sheet PLS_INTEGER := NULL
    );
--
  PROCEDURE hyperlink
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_url VARCHAR2
    , p_value VARCHAR2 := NULL
    , p_sheet PLS_INTEGER := NULL
    );
--
  PROCEDURE comment
    ( p_col PLS_INTEGER
    , p_row PLS_INTEGER
    , p_text VARCHAR2
    , p_author VARCHAR2 := NULL
    , p_width PLS_INTEGER := 150  -- pixels
    , p_height PLS_INTEGER := 100  -- pixels
    , p_sheet PLS_INTEGER := NULL
    );
--
  PROCEDURE mergecells
    ( p_tl_col PLS_INTEGER -- top left
    , p_tl_row PLS_INTEGER
    , p_br_col PLS_INTEGER -- bottom right
    , p_br_row PLS_INTEGER
    , p_sheet PLS_INTEGER := NULL
    );
--
  PROCEDURE list_validation
    ( p_sqref_col PLS_INTEGER
    , p_sqref_row PLS_INTEGER
    , p_tl_col PLS_INTEGER -- top left
    , p_tl_row PLS_INTEGER
    , p_br_col PLS_INTEGER -- bottom right
    , p_br_row PLS_INTEGER
    , p_style VARCHAR2 := 'stop' -- stop, warning, information
    , p_title VARCHAR2 := NULL
    , p_prompt VARCHAR2 := NULL
    , p_show_error boolean := false
    , p_error_title VARCHAR2 := NULL
    , p_error_txt VARCHAR2 := NULL
    , p_sheet PLS_INTEGER := NULL
    );
--
  PROCEDURE list_validation ( p_sqref_col PLS_INTEGER
                            , p_sqref_row PLS_INTEGER
                            , p_defined_name VARCHAR2
                            , p_style VARCHAR2 := 'stop' -- stop, warning, information
                            , p_title VARCHAR2 := NULL
                            , p_prompt VARCHAR2 := NULL
                            , p_show_error boolean := FALSE
                            , p_error_title VARCHAR2 := NULL
                            , p_error_txt VARCHAR2 := NULL
                            , p_sheet PLS_INTEGER := NULL
                            )
  ;

  PROCEDURE defined_name ( p_tl_col PLS_INTEGER -- top left
                         , p_tl_row PLS_INTEGER
                         , p_br_col PLS_INTEGER -- bottom right
                         , p_br_row PLS_INTEGER
                         , p_name VARCHAR2
                         , p_sheet PLS_INTEGER := NULL
                         , p_localsheet PLS_INTEGER := NULL
                         )
  ;

  PROCEDURE set_column_width ( p_col PLS_INTEGER
                             , p_width NUMBER
                             , p_sheet PLS_INTEGER := NULL
                             )
  ;

  PROCEDURE set_column ( p_col PLS_INTEGER
                       , p_numFmtId PLS_INTEGER := NULL
                       , p_fontId PLS_INTEGER := NULL
                       , p_fillId PLS_INTEGER := NULL
                       , p_borderId PLS_INTEGER := NULL
                       , p_alignment tp_alignment := NULL
                       , p_sheet PLS_INTEGER := NULL
                       )
  ;

  PROCEDURE set_row ( p_row PLS_INTEGER
                    , p_numFmtId PLS_INTEGER := NULL
                    , p_fontId PLS_INTEGER := NULL
                    , p_fillId PLS_INTEGER := NULL
                    , p_borderId PLS_INTEGER := NULL
                    , p_alignment tp_alignment := NULL
                    , p_sheet PLS_INTEGER := NULL
                    )
  ;

  PROCEDURE freeze_rows ( p_nr_rows PLS_INTEGER := 1
                        , p_sheet PLS_INTEGER := NULL
                        )
  ;

  PROCEDURE freeze_cols ( p_nr_cols PLS_INTEGER := 1
                        , p_sheet PLS_INTEGER := NULL
                        )
  ;

  PROCEDURE freeze_pane ( p_col PLS_INTEGER
                        , p_row PLS_INTEGER
                        , p_sheet PLS_INTEGER := NULL
                        )
  ;

  PROCEDURE set_autofilter ( p_column_start PLS_INTEGER := NULL
                           , p_column_end PLS_INTEGER := NULL
                           , p_row_start PLS_INTEGER := NULL
                           , p_row_end PLS_INTEGER := NULL
                           , p_sheet PLS_INTEGER := NULL
                           )
  ;

  FUNCTION FINISH
    RETURN BLOB
  ;

  FUNCTION query2sheet ( p_sql VARCHAR2
                       , p_column_headers boolean := TRUE
                       , p_directory VARCHAR2 := NULL
                       , p_filename VARCHAR2 := NULL
                       , p_sheet PLS_INTEGER := NULL
                       )
    RETURN BLOB
  ;


END;

/
