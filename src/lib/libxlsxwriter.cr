module CrXLSXWriter
  @[Link("xlsxwriter")]
  lib LibXLSXWriter
    type Workbook = Void*
    type Worksheet = Void*
    type Chartsheet = Void*
    type Format = Void*
    type Chart = Void*
    type ChartSeries = Void*
    type SeriesErrorBars = Void*
    type ChartAxis = Void*

    alias Str = UInt8*

    LXW_DEF_ROW_HEIGHT = 15

    alias Row = UInt32
    alias Col = UInt16

    alias Color = Int32

    enum LXWError
      LXW_NO_ERROR
      LXW_ERROR_MEMORY_MALLOC_FAILED
      LXW_ERROR_CREATING_XLSX_FILE
      LXW_ERROR_CREATING_TMPFILE
      LXW_ERROR_READING_TMPFILE
      LXW_ERROR_ZIP_FILE_OPERATION
      LXW_ERROR_ZIP_PARAMETER_ERROR
      LXW_ERROR_ZIP_BAD_ZIP_FILE
      LXW_ERROR_ZIP_INTERNAL_ERROR
      LXW_ERROR_ZIP_FILE_ADD
      LXW_ERROR_ZIP_CLOSE
      LXW_ERROR_NULL_PARAMETER_IGNORED
      LXW_ERROR_PARAMETER_VALIDATION
      LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED
      LXW_ERROR_INVALID_SHEETNAME_CHARACTER
      LXW_ERROR_SHEETNAME_START_END_APOSTROPHE
      LXW_ERROR_SHEETNAME_ALREADY_USED
      LXW_ERROR_SHEETNAME_RESERVED
      LXW_ERROR_32_STRING_LENGTH_EXCEEDED
      LXW_ERROR_128_STRING_LENGTH_EXCEEDED
      LXW_ERROR_255_STRING_LENGTH_EXCEEDED
      LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED
      LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND
      LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE
      LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED
      LXW_ERROR_IMAGE_DIMENSIONS
    end

    struct DocProperties
      title : Str
      subject : Str
      author : Str
      manager : Str
      company : Str
      category : Str
      keywords : Str
      comments : Str
      status : Str
      hyperlink_base : Str
    end

    enum UnderlineStyle : Int8
      LXW_UNDERLINE_SINGLE            = 1
      LXW_UNDERLINE_DOUBLE            = 2
      LXW_UNDERLINE_SINGLE_ACCOUNTING = 3
      LXW_UNDERLINE_DOUBLE_ACCOUNTING = 4
    end

    enum FontScript : Int8
      LXW_FONT_SUPERSCRIPT = 1
      LXW_FONT_SUBSCRIPT   = 2
    end

    enum Alignment : Int8
      LXW_ALIGN_NONE
      LXW_ALIGN_LEFT
      LXW_ALIGN_CENTER
      LXW_ALIGN_RIGHT
      LXW_ALIGN_FILL
      LXW_ALIGN_JUSTIFY
      LXW_ALIGN_CENTER_ACROSS
      LXW_ALIGN_DISTRIBUTED
      LXW_ALIGN_VERTICAL_TOP
      LXW_ALIGN_VERTICAL_BOTTOM
      LXW_ALIGN_VERTICAL_JUSTIFY
      LXW_ALIGN_VERTICAL_DISTRIBUTED
    end

    enum Border : Int8
      LXW_BORDER_NONE
      LXW_BORDER_THIN
      LXW_BORDER_MEDIUM
      LXW_BORDER_DASHED
      LXW_BORDER_DOTTED
      LXW_BORDER_THICK
      LXW_BORDER_DOUBLE
      LXW_BORDER_HAIR
      LXW_BORDER_MEDIUM_DASHED
      LXW_BORDER_DASH_DOT
      LXW_BORDER_MEDIUM_DASH_DOT
      LXW_BORDER_DASH_DOT_DOT
      LXW_BORDER_MEDIUM_DASH_DOT_DOT
      LXW_BORDER_SLANT_DASH_DOT
    end

    enum Pattern : Int8
      LXW_PATTERN_NONE
      LXW_PATTERN_SOLID
      LXW_PATTERN_MEDIUM_GRAY
      LXW_PATTERN_DARK_GRAY
      LXW_PATTERN_LIGHT_GRAY
      LXW_PATTERN_DARK_HORIZONTAL
      LXW_PATTERN_DARK_VERTICAL
      LXW_PATTERN_DARK_DOWN
      LXW_PATTERN_DARK_UP
      LXW_PATTERN_DARK_GRID
      LXW_PATTERN_DARK_TRELLIS
      LXW_PATTERN_LIGHT_HORIZONTAL
      LXW_PATTERN_LIGHT_VERTICAL
      LXW_PATTERN_LIGHT_DOWN
      LXW_PATTERN_LIGHT_UP
      LXW_PATTERN_LIGHT_GRID
      LXW_PATTERN_LIGHT_TRELLIS
      LXW_PATTERN_GRAY_125
      LXW_PATTERN_GRAY_0625
    end

    enum ChartType : Int8
      LXW_CHART_NONE
      LXW_CHART_AREA
      LXW_CHART_AREA_STACKED
      LXW_CHART_AREA_STACKED_PERCENT
      LXW_CHART_BAR
      LXW_CHART_BAR_STACKED
      LXW_CHART_BAR_STACKED_PERCENT
      LXW_CHART_COLUMN
      LXW_CHART_COLUMN_STACKED
      LXW_CHART_COLUMN_STACKED_PERCENT
      LXW_CHART_DOUGHNUT
      LXW_CHART_LINE
      LXW_CHART_PIE
      LXW_CHART_SCATTER
      LXW_CHART_SCATTER_STRAIGHT
      LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS
      LXW_CHART_SCATTER_SMOOTH
      LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS
      LXW_CHART_RADAR
      LXW_CHART_RADAR_WITH_MARKERS
      LXW_CHART_RADAR_FILLED
    end

    enum ChartLegendPosition : Int8
      LXW_CHART_LEGEND_NONE
      LXW_CHART_LEGEND_RIGHT
      LXW_CHART_LEGEND_LEFT
      LXW_CHART_LEGEND_TOP
      LXW_CHART_LEGEND_BOTTOM
      LXW_CHART_LEGEND_TOP_RIGHT
      LXW_CHART_LEGEND_OVERLAY_RIGHT
      LXW_CHART_LEGEND_OVERLAY_LEFT
      LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT
    end

    enum ChartLineDashType : Int8
      LXW_CHART_LINE_DASH_SOLID
      LXW_CHART_LINE_DASH_ROUND_DOT
      LXW_CHART_LINE_DASH_SQUARE_DOT
      LXW_CHART_LINE_DASH_DASH
      LXW_CHART_LINE_DASH_DASH_DOT
      LXW_CHART_LINE_DASH_LONG_DASH
      LXW_CHART_LINE_DASH_LONG_DASH_DOT
      LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT
    end

    enum ChartMarkerType : Int8
      LXW_CHART_MARKER_AUTOMATIC
      LXW_CHART_MARKER_NONE
      LXW_CHART_MARKER_SQUARE
      LXW_CHART_MARKER_DIAMOND
      LXW_CHART_MARKER_TRIANGLE
      LXW_CHART_MARKER_X
      LXW_CHART_MARKER_STAR
      LXW_CHART_MARKER_SHORT_DASH
      LXW_CHART_MARKER_LONG_DASH
      LXW_CHART_MARKER_CIRCLE
      LXW_CHART_MARKER_PLUS
    end

    enum ChartPatternType : Int8
      LXW_CHART_PATTERN_NONE
      LXW_CHART_PATTERN_PERCENT_5
      LXW_CHART_PATTERN_PERCENT_10
      LXW_CHART_PATTERN_PERCENT_20
      LXW_CHART_PATTERN_PERCENT_25
      LXW_CHART_PATTERN_PERCENT_30
      LXW_CHART_PATTERN_PERCENT_40
      LXW_CHART_PATTERN_PERCENT_50
      LXW_CHART_PATTERN_PERCENT_60
      LXW_CHART_PATTERN_PERCENT_70
      LXW_CHART_PATTERN_PERCENT_75
      LXW_CHART_PATTERN_PERCENT_80
      LXW_CHART_PATTERN_PERCENT_90
      LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL
      LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL
      LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL
      LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL
      LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL
      LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL
      LXW_CHART_PATTERN_LIGHT_VERTICAL
      LXW_CHART_PATTERN_LIGHT_HORIZONTAL
      LXW_CHART_PATTERN_NARROW_VERTICAL
      LXW_CHART_PATTERN_NARROW_HORIZONTAL
      LXW_CHART_PATTERN_DARK_VERTICAL
      LXW_CHART_PATTERN_DARK_HORIZONTAL
      LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL
      LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL
      LXW_CHART_PATTERN_DASHED_HORIZONTAL
      LXW_CHART_PATTERN_DASHED_VERTICAL
      LXW_CHART_PATTERN_SMALL_CONFETTI
      LXW_CHART_PATTERN_LARGE_CONFETTI
      LXW_CHART_PATTERN_ZIGZAG
      LXW_CHART_PATTERN_WAVE
      LXW_CHART_PATTERN_DIAGONAL_BRICK
      LXW_CHART_PATTERN_HORIZONTAL_BRICK
      LXW_CHART_PATTERN_WEAVE
      LXW_CHART_PATTERN_PLAID
      LXW_CHART_PATTERN_DIVOT
      LXW_CHART_PATTERN_DOTTED_GRID
      LXW_CHART_PATTERN_DOTTED_DIAMOND
      LXW_CHART_PATTERN_SHINGLE
      LXW_CHART_PATTERN_TRELLIS
      LXW_CHART_PATTERN_SPHERE
      LXW_CHART_PATTERN_SMALL_GRID
      LXW_CHART_PATTERN_LARGE_GRID
      LXW_CHART_PATTERN_SMALL_CHECK
      LXW_CHART_PATTERN_LARGE_CHECK
      LXW_CHART_PATTERN_OUTLINED_DIAMOND
      LXW_CHART_PATTERN_SOLID_DIAMOND
    end

    enum ChartLabelPosition : Int8
      LXW_CHART_LABEL_POSITION_DEFAULT
      LXW_CHART_LABEL_POSITION_CENTER
      LXW_CHART_LABEL_POSITION_RIGHT
      LXW_CHART_LABEL_POSITION_LEFT
      LXW_CHART_LABEL_POSITION_ABOVE
      LXW_CHART_LABEL_POSITION_BELOW
      LXW_CHART_LABEL_POSITION_INSIDE_BASE
      LXW_CHART_LABEL_POSITION_INSIDE_END
      LXW_CHART_LABEL_POSITION_OUTSIDE_END
      LXW_CHART_LABEL_POSITION_BEST_FIT
    end

    enum ChartLabelSeperator : Int8
      LXW_CHART_LABEL_SEPARATOR_COMMA
      LXW_CHART_LABEL_SEPARATOR_SEMICOLON
      LXW_CHART_LABEL_SEPARATOR_PERIOD
      LXW_CHART_LABEL_SEPARATOR_NEWLINE
      LXW_CHART_LABEL_SEPARATOR_SPACE
    end

    enum ChartAxisType : Int8
      LXW_CHART_AXIS_TYPE_X
      LXW_CHART_AXIS_TYPE_Y
    end

    enum ChartAxisTickPosition : Int8
      LXW_CHART_AXIS_POSITION_ON_TICK
      LXW_CHART_AXIS_POSITION_BETWEEN
    end

    enum ChartAxisLabelPosition : Int8
      LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO
      LXW_CHART_AXIS_LABEL_POSITION_HIGH
      LXW_CHART_AXIS_LABEL_POSITION_LOW
      LXW_CHART_AXIS_LABEL_POSITION_NONE
    end

    enum ChartAxisLabelAlignment : Int8
      LXW_CHART_AXIS_LABEL_ALIGN_CENTER
      LXW_CHART_AXIS_LABEL_ALIGN_LEFT
      LXW_CHART_AXIS_LABEL_ALIGN_RIGHT
    end

    enum ChartAxisDisplayUnit : Int8
      LXW_CHART_AXIS_UNITS_NONE
      LXW_CHART_AXIS_UNITS_HUNDREDS
      LXW_CHART_AXIS_UNITS_THOUSANDS
      LXW_CHART_AXIS_UNITS_TEN_THOUSANDS
      LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS
      LXW_CHART_AXIS_UNITS_MILLIONS
      LXW_CHART_AXIS_UNITS_TEN_MILLIONS
      LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS
      LXW_CHART_AXIS_UNITS_BILLIONS
      LXW_CHART_AXIS_UNITS_TRILLIONS
    end

    enum ChartAxisTickMark : Int8
      LXW_CHART_AXIS_TICK_MARK_DEFAULT
      LXW_CHART_AXIS_TICK_MARK_NONE
      LXW_CHART_AXIS_TICK_MARK_INSIDE
      LXW_CHART_AXIS_TICK_MARK_OUTSIDE
      LXW_CHART_AXIS_TICK_MARK_CROSSING
    end

    enum ChartBlank : Int8
      LXW_CHART_BLANKS_AS_GAP
      LXW_CHART_BLANKS_AS_ZERO
      LXW_CHART_BLANKS_AS_CONNECTED
    end

    enum ChartErrorBarType : Int8
      LXW_CHART_ERROR_BAR_TYPE_STD_ERROR
      LXW_CHART_ERROR_BAR_TYPE_FIXED
      LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE
      LXW_CHART_ERROR_BAR_TYPE_STD_DEV
    end

    enum ChartErrorBarAxis : Int8
      LXW_CHART_ERROR_BAR_AXIS_X
      LXW_CHART_ERROR_BAR_AXIS_Y
    end

    enum ChartErrorBarCap : Int8
      LXW_CHART_ERROR_BAR_END_CAP
      LXW_CHART_ERROR_BAR_NO_CAP
    end

    enum ChartTrendlineType : Int8
      LXW_CHART_TRENDLINE_TYPE_LINEAR
      LXW_CHART_TRENDLINE_TYPE_LOG
      LXW_CHART_TRENDLINE_TYPE_POLY
      LXW_CHART_TRENDLINE_TYPE_POWER
      LXW_CHART_TRENDLINE_TYPE_EXP
      LXW_CHART_TRENDLINE_TYPE_AVERAGE
    end

    enum ChartErrorBarDirection : Int8
      LXW_CHART_ERROR_BAR_DIR_BOTH
      LXW_CHART_ERROR_BAR_DIR_PLUS
      LXW_CHART_ERROR_BAR_DIR_MINUS
    end

    enum Gridlines : UInt8
      LXW_HIDE_ALL_GRIDLINES 	
      LXW_SHOW_SCREEN_GRIDLINES 	
      LXW_SHOW_PRINT_GRIDLINES 	
      LXW_SHOW_ALL_GRIDLINES 
    end

    struct ChartLine
      color : Color
      none : Bool
      width : LibC::Float
      dash_type : ChartLineDashType
      transparency : UInt8
    end

    struct ChartFill
      color : Color
      none : Bool
      transparency : UInt8
    end

    struct ChartPattern
      fg_color : Color
      bg_color : Color
      type : ChartPatternType
    end

    struct ChartFont
      name : Str
      size : LibC::Double
      bold : Bool
      italic : Bool
      underline : Bool
      rotation : Int32
      color : Color
      pitch_family : UInt8
      charset : UInt8
      baseline : Int8
    end

    struct Datetime
      year : Int32
      month : Int32
      day : Int32
      hour : Int32
      min : Int32
      sec : LibC::Double
    end

    struct RichString
      format : Format*
      string : Str
    end

    struct RowColOptions
      hidden : Bool
      level : Bool
      collapsed : Bool
    end

    struct ImageOptions
      x_offset : Int32
      y_offset : Int32
      x_scale : LibC::Double
      y_scale : LibC::Double
    end

    struct Protection
      no_select_locked_cells : Bool
      no_select_unlocked_cells : Bool
      format_cells : Bool
      format_columns : Bool
      format_rows : Bool
      insert_columns : Bool
      insert_rows : Bool
      insert_hyperlinks : Bool
      delete_columns : Bool
      delete_rows : Bool
      sort : Bool
      autofilter : Bool
      pivot_tables : Bool
      scenarios : Bool
      objects : Bool
      no_content : Bool
      no_objects : Bool
    end

    struct HeaderFooterOptions
      margin : LibC::Double
    end

    # struct DataValidation
    # end

    fun workbook_new(path : Str) : Workbook*
    fun workbook_add_worksheet(workbook : Workbook*, sheetname : Str) : Worksheet*
    fun workbook_close(workbook : Workbook*) : LXWError

    fun workbook_set_properties(workbook : Workbook*, properties : DocProperties*) : LXWError
    fun workbook_set_custom_property_string(workbook : Workbook*, name : Str, value : Str) : LXWError
    fun workbook_set_custom_property_number(workbook : Workbook*, name : Str, value : LibC::Double) : LXWError
    fun workbook_set_custom_property_boolean(workbook : Workbook*, name : Str, value : Bool) : LXWError
    fun workbook_set_custom_property_datetime(workbook : Workbook*, name : Str, value : Datetime*) : LXWError
    fun workbook_define_name(workbook : Workbook*, name : Str, formula : Str) : LXWError
    fun workbook_add_vba_project(workbook : Workbook*, filename : Str) : LXWError
    fun workbook_set_vba_name(workbook : Workbook*, name : Str) : LXWError

    fun workbook_get_worksheet_by_name(workbook : Workbook*, name : Str) : Worksheet*
    fun workbook_get_chartsheet_by_name(workbook : Workbook*, name : Str) : Chartsheet*

    fun workbook_add_format(workbook : Workbook*) : Format*
    fun workbook_add_chart(workbook : Workbook*, chart_type : ChartType) : Chart*

    fun worksheet_write_string(worksheet : Worksheet*, row : Row, col : Col, string : Str, format : Format*) : LXWError
    fun worksheet_write_number(worksheet : Worksheet*, row : Row, col : Col, value : LibC::Double, format : Format*) : LXWError
    fun worksheet_write_formula(worksheet : Worksheet*, row : Row, col : Col, value : Str, format : Format*) : LXWError
    # workbook_write_array_formula
    fun worksheet_write_datetime(worksheet : Worksheet*, row : Row, col : Col, datetime : Datetime*, format : Format*) : LXWError
    fun worksheet_write_url(worksheet : Worksheet*, row : Row, col : Col, url : Str, format : Format*) : LXWError
    # worksheet_write_boolean
    fun worksheet_write_blank(worksheet : Worksheet*, row : Row, col : Col, format : Format*) : LXWError
    fun worksheet_write_formula_num(worksheet : Worksheet*, row : Row, col : Col, formula : Str, format : Format*, result : LibC::Double) : LXWError
    fun worksheet_write_rich_string(worksheet : Worksheet*, row : Row, col : Col, rich_string : StaticArray(RichString*, 16), format : Format*) : LXWError

    fun worksheet_set_row(worksheet : Worksheet*, row : Row, height : LibC::Double, format : Format*) : LXWError
    fun worksheet_set_row_opt(worksheet : Worksheet*, row : Row, height : LibC::Double, format : Format*, options : RowColOptions*) : LXWError
    fun worksheet_set_column(worksheet : Worksheet*, first_col : Col, last_col : Col, width : LibC::Double, format : Format*) : LXWError
    fun worksheet_set_column_opt(worksheet : Worksheet*, first_col : Col, last_col : Col, width : LibC::Double, format : Format*, options : RowColOptions*) : LXWError

    fun worksheet_insert_image(worksheet : Worksheet*, row : Row, col : Col, filename : Str) : LXWError
    fun worksheet_insert_image_opt(worksheet : Worksheet*, row : Row, col : Col, filename : Str, options : ImageOptions*) : LXWError
    fun worksheet_insert_chart(worksheet : Worksheet*, row : Row, col : Col, chart : Chart*) : LXWError
    fun worksheet_insert_chart_opt(worksheet : Worksheet*, row : Row, col : Col, chart : Chart*, user_options : ImageOptions*) : LXWError

    fun worksheet_merge_range(
      worksheet : Worksheet*,
      first_row : Row,
      first_col : Col, 
      last_row : Row,
      last_col : Col,
      string : Str,
      format : Format*,
    ) : LXWError
    fun worksheet_autofilter(
      worksheet : Worksheet*,
      first_row : Row,
      first_col : Col, 
      last_row : Row,
      last_col : Col,
    ) : LXWError
    # fun worksheet_data_validation_cell
    fun worksheet_activate(worksheet : Worksheet*) : Void
    fun worksheet_select(worksheet : Worksheet*) : Void
    fun worksheet_hide(worksheet : Worksheet*) : Void
    fun worksheet_set_first_sheet(worksheet : Worksheet*) : Void
    fun worksheet_set_first_sheet(worksheet : Worksheet*) : Void
    fun worksheet_freeze_panes(worksheet : Worksheet*, row : Row, col : Col) : Void
    fun worksheet_split_panes(worksheet : Worksheet*, vertical : LibC::Double, horizontal : LibC::Double) : Void
    fun worksheet_set_selection(
      worksheet : Worksheet*,
      first_row : Row,
      first_col : Col, 
      last_row : Row,
      last_col : Col,
    ) : Void
    fun worksheet_set_landscape(worksheet : Worksheet*) : Void
    fun worksheet_set_portrait(worksheet : Worksheet*) : Void
    fun worksheet_set_page_view(worksheet : Worksheet*) : Void
    fun worksheet_set_paper(worksheet : Worksheet*, paper_type : UInt8) : Void
    fun worksheet_set_margins(
      worksheet : Worksheet*,
      left : LibC::Double,
      right : LibC::Double,
      top : LibC::Double,
      bottom : LibC::Double
    ) : Void
    fun worksheet_set_header(worksheet : Worksheet*, string : Str) : LXWError
    fun worksheet_set_header_opt(worksheet : Worksheet*, string : Str, options : HeaderFooterOptions*) : LXWError
    fun worksheet_set_footer(worksheet : Worksheet*, string : Str) : LXWError
    fun worksheet_set_footer_opt(worksheet : Worksheet*, string : Str, options : HeaderFooterOptions*) : LXWError
    fun worksheet_set_h_pagebreaks(worksheet : Worksheet*, breaks : Row[16]) : LXWError
    fun worksheet_set_v_pagebreaks(worksheet : Worksheet*, breaks : Col[16]) : LXWError
    fun worksheet_print_across(worksheet : Worksheet*) : Void
    fun worksheet_set_zoom(worksheet : Worksheet*, scale : UInt16) : Void
    fun worksheet_gridlines(worksheet : Worksheet*, option : Gridlines) : Void
    fun worksheet_center_horizontally(worksheet : Worksheet*) : Void
    fun worksheet_center_vertically(worksheet : Worksheet*) : Void
    fun worksheet_print_row_col_headers(worksheet : Worksheet*) : Void
    fun worksheet_repeat_rows(worksheet : Worksheet*, first_row : Row, last_row : Row) : LXWError
    fun worksheet_repeat_columns(worksheet : Worksheet*, first_row : Col, last_row : Col) : LXWError
    fun worksheet_print_area(
      worksheet : Worksheet*,
      first_row : Row,
      first_col : Col, 
      last_row : Row,
      last_col : Col
    ) : LXWError
    fun worksheet_fit_to_pages(worksheet : Worksheet*, width : UInt16, height : UInt16) : Void
    fun worksheet_set_start_page(worksheet : Worksheet*, start_page : UInt16) : Void
    fun worksheet_set_print_scale(worksheet : Worksheet*, scale : UInt16) : Void
    fun worksheet_right_to_left(worksheet : Worksheet*) : Void
    fun worksheet_hide_zero(worksheet : Worksheet*) : Void
    fun worksheet_set_tab_color(worksheet : Worksheet*, color : Color) : Void
    fun worksheet_protect(worksheet : Worksheet*, password : Str, options : Protection*) : Void
    fun worksheet_outline_settings(
      worksheet : Worksheet*,
      visible : Bool,
      symbols_below : Bool,
      symbols_right : Bool,
      auto_style : Bool
    ) : Void
    fun worksheet_set_default_row(worksheet : Worksheet*, height : LibC::Double, hide_unused_rows : Bool) : Void
    fun worksheet_set_vba_name(worksheet : Worksheet*, name : Str) : LXWError

    # format.h
    fun format_set_bold(format : Format*) : Void
    fun format_set_font_color(format : Format*, color : Color) : Void
    fun format_set_font_name(format : Format*, font_name : Str) : Void
    fun format_set_font_size(format : Format*, font_size : LibC::Double) : Void
    fun format_set_italic(format : Format*) : Void
    fun format_set_underline(format : Format*, style : UnderlineStyle) : Void
    fun format_set_font_strikeout(format : Format*) : Void
    fun format_set_font_script(format : Format*, script : FontScript) : Void
    fun format_set_num_format(format : Format*, num_format : Str) : Void
    fun format_set_unlocked(format : Format*) : Void
    fun format_set_hidden(format : Format*) : Void
    fun format_set_align(format : Format*, align : Alignment) : Void
    fun format_set_text_wrap(format : Format*) : Void
    fun format_set_rotation(format : Format*, angle : Int16) : Void
    fun format_set_indent(format : Format*, level : UInt8) : Void
    fun format_set_pattern(format : Format*, pattern : Pattern) : Void
    fun format_set_bg_color(format : Format*, color : Color) : Void
    fun format_set_fg_color(format : Format*, color : Color) : Void
    fun format_set_fg_color(format : Format*, color : Color) : Void
    fun format_set_border(format : Format*, style : Border) : Void
    fun format_set_bottom(format : Format*, style : Border) : Void
    fun format_set_top(format : Format*, style : Border) : Void
    fun format_set_left(format : Format*, style : Border) : Void
    fun format_set_right(format : Format*, style : Border) : Void
    fun format_set_border_color(format : Format*, color : Color) : Void
    fun format_set_bottom_color(format : Format*, color : Color) : Void
    fun format_set_top_color(format : Format*, color : Color) : Void
    fun format_set_left_color(format : Format*, color : Color) : Void
    fun format_set_right_color(format : Format*, color : Color) : Void

    # chart.h
    fun chart_add_series(chart : Chart*, categories : Str, values : Str) : ChartSeries*
    fun chart_series_set_values(series : ChartSeries*, sheetname : Str, first_row : Row, first_col : Col, last_row : Row, last_col : Col) : Void
    fun chart_series_set_categories(series : ChartSeries*, sheetname : Str, first_row : Row, first_col : Col, last_row : Row, last_col : Col) : Void
    fun chart_series_set_name(series : ChartSeries*, name : Str) : Void
    fun chart_series_set_name_range(series : ChartSeries*, sheetname : Str, row : Row, col : Col) : Void
    fun chart_series_set_line(series : ChartSeries*, line : ChartLine*) : Void
    fun chart_series_set_fill(series : ChartSeries*, fill : ChartFill*) : Void
    fun chart_series_set_invert_if_negative(series : ChartSeries*) : Void
    fun chart_series_set_pattern(series : ChartSeries*, pattern : ChartPattern) : Void
    fun chart_series_set_marker_type(series : ChartSeries*, type : ChartMarkerType) : Void
    fun chart_series_set_marker_size(series : ChartSeries*, size : UInt8) : Void
    fun chart_series_set_marker_line(series : ChartSeries*, line : ChartLine*) : Void
    fun chart_series_set_marker_fill(series : ChartSeries*, fill : ChartFill*) : Void
    fun chart_series_set_marker_pattern(series : ChartSeries*, pattern : ChartPattern) : Void
    fun chart_series_set_smooth(series : ChartSeries*, smooth : Bool) : Void
    fun chart_series_set_labels(series : ChartSeries*) : Void
    fun chart_series_set_labels_options(series : ChartSeries*, show_name : Bool, show_category : Bool, show_value : Bool) : Void
    fun chart_series_set_labels_separator(series : ChartSeries*, separator : ChartLabelSeperator) : Void
    fun chart_series_set_labels_position(series : ChartSeries*, position : ChartLabelPosition) : Void
    fun chart_series_set_labels_leader_line(series : ChartSeries*) : Void
    fun chart_series_set_labels_legend(series : ChartSeries*) : Void
    fun chart_series_set_labels_percentage(series : ChartSeries*) : Void
    fun chart_series_set_labels_num_format(series : ChartSeries*, num_format : Str) : Void
    fun chart_series_set_labels_font(series : ChartSeries*, font : ChartFont*) : Void
    fun chart_series_set_trendline(series : ChartSeries*, type : ChartTrendlineType, value : UInt8) : Void
    fun chart_series_set_trendline_forecast(series : ChartSeries*, forward : LibC::Double, backward : LibC::Double) : Void
    fun chart_series_set_trendline_equation(series : ChartSeries*) : Void
    fun chart_series_set_trendline_r_squared(series : ChartSeries*) : Void
    fun chart_series_set_trendline_intercept(series : ChartSeries*, intercept : LibC::Double) : Void
    fun chart_series_set_trendline_name(series : ChartSeries*, name : Str) : Void
    fun chart_series_set_trendline_line(series : ChartSeries*, line : ChartLine*) : Void
    fun chart_series_get_error_bars(series : ChartSeries*, axis_type : ChartErrorBarAxis) : SeriesErrorBars*

    fun chart_series_set_error_bars(error_bars : SeriesErrorBars*, type : ChartErrorBarType, value : LibC::Double) : Void
    fun chart_series_set_error_bars_direction(error_bars : SeriesErrorBars*, direction : ChartErrorBarDirection) : Void
    fun chart_series_set_error_bars_endcap(error_bars : SeriesErrorBars*, endcap : ChartErrorBarCap) : Void
    fun chart_series_set_error_bars_line(error_bars : SeriesErrorBars*, line : ChartLine*) : Void

    fun chart_axis_get(chart : Chart*, axis_type : ChartAxisType) : ChartAxis*
    fun chart_axis_set_name(axis : ChartAxis*, name : Str) : Void
    fun chart_axis_set_name_range(axis : ChartAxis*, sheetname : Str, row : Row, col : Col) : Void
    fun chart_axis_set_name_font(axis : ChartAxis*, font : ChartFont*) : Void
    fun chart_axis_set_num_font(axis : ChartAxis*, font : ChartFont*) : Void
    fun chart_axis_set_num_format(axis : ChartAxis*, num_format : Str) : Void
    fun chart_axis_set_line(axis : ChartAxis*, chart_line : ChartLine*) : Void
    fun chart_axis_set_fill(axis : ChartAxis*, fill : ChartFill*) : Void
    fun chart_axis_set_pattern(axis : ChartAxis*, pattern : ChartPattern*) : Void
    fun chart_axis_set_reverse(axis : ChartAxis*) : Void
    fun chart_axis_set_crossing(axis : ChartAxis*, value : LibC::Double) : Void
    fun chart_axis_set_crossing_max(axis : ChartAxis*) : Void
    fun chart_axis_off(axis : ChartAxis*) : Void
    fun chart_axis_set_position(axis : ChartAxis*, position : ChartAxisTickPosition) : Void
    fun chart_axis_set_label_position(axis : ChartAxis*, position : ChartAxisLabelPosition) : Void
    fun chart_axis_set_label_align(axis : ChartAxis*, align : ChartAxisLabelAlignment) : Void
    fun chart_axis_set_min(axis : ChartAxis*, min : LibC::Double) : Void
    fun chart_axis_set_max(axis : ChartAxis*, max : LibC::Double) : Void
    fun chart_axis_set_log_base(axis : ChartAxis*, log_base : UInt16) : Void
    fun chart_axis_set_major_tick_mark(axis : ChartAxis*, type : ChartAxisTickMark) : Void
    fun chart_axis_set_minor_tick_mark(axis : ChartAxis*, type : ChartAxisTickMark) : Void
    fun chart_axis_set_interval_unit(axis : ChartAxis*, unit : UInt16) : Void
    fun chart_axis_set_interval_tick(axis : ChartAxis*, unit : UInt16) : Void
    fun chart_axis_set_major_unit(axis : ChartAxis*, unit : LibC::Double) : Void
    fun chart_axis_set_minor_unit(axis : ChartAxis*, unit : LibC::Double) : Void
    fun chart_axis_set_display_units(axis : ChartAxis*, units : ChartAxisDisplayUnit) : Void
    fun chart_axis_set_display_units_visible(axis : ChartAxis*, visible : Bool) : Void
    fun chart_axis_major_gridlines_set_visible(axis : ChartAxis*, visible : Bool) : Void
    fun chart_axis_minor_gridlines_set_visible(axis : ChartAxis*, visible : Bool) : Void
    fun chart_axis_major_gridlines_set_line(axis : ChartAxis*, line : ChartLine*) : Void
    fun chart_axis_minor_gridlines_set_line(axis : ChartAxis*, line : ChartLine*) : Void

    fun chart_title_set_name(chart : Chart*, name : Str) : Void
    fun chart_title_set_name_range(chart : Chart*, sheetname : Str, row : Row, col : Col) : Void
    fun chart_title_set_name_font(chart : Chart*, font : ChartFont*) : Void
    fun chart_title_off(chart : Chart*) : Void

    fun chart_legend_set_position(chart : Chart*, position : ChartLegendPosition) : Void
    fun chart_legend_set_font(chart : Chart*, font : ChartFont*) : Void
    fun chart_legend_delete_series(chart : Chart*, delete_series : Int16[16]) : LXWError

    fun chart_chartarea_set_line(chart : Chart*, line : ChartLine*) : Void
    fun chart_chartarea_set_fill(chart : Chart*, line : ChartFill*) : Void
    fun chart_chartarea_set_pattern(chart : Chart*, line : ChartPattern*) : Void

    fun chart_plotarea_set_line(chart : Chart*, line : ChartLine*) : Void
    fun chart_plotarea_set_fill(chart : Chart*, line : ChartFill*) : Void
    fun chart_plotarea_set_pattern(chart : Chart*, line : ChartPattern*) : Void

    fun chart_set_style(chart : Chart*, style_id : UInt8) : Void
    fun chart_set_table(chart : Chart*) : Void
    fun chart_set_table_grid(chart : Chart*, horizontal : Bool, vertical : Bool, outline : Bool, legend_keys : Bool) : Void
    fun chart_set_up_down_bars(chart : Chart*) : Void
    fun chart_set_up_down_bars_format(chart : Chart*, up_bar_line : ChartLine*, up_bar_fill : ChartFill*, down_bar_line : ChartLine*, down_bar_fill : ChartFill*) : Void
    fun chart_set_drop_lines(chart : Chart*, line : ChartLine*) : Void
    fun chart_set_high_low_lines(chart : Chart*, line : ChartLine*) : Void
    fun chart_set_series_overlap(chart : Chart*, overlap : Int8) : Void
    fun chart_set_series_gap(chart : Chart*, gap : UInt16) : Void
    fun chart_show_blanks_as(chart : Chart*, option : ChartBlank) : Void
    fun chart_set_rotation(chart : Chart*, rotation : UInt16) : Void
    fun chart_set_hole_size(chart : Chart*, size : UInt8) : Void

    # chartsheet.h
    fun chartsheet_set_chart(chartsheet : Chartsheet*, chart : Chart*) : LXWError
    fun chartsheet_activate(chartsheet : Chartsheet*) : Void
    fun chartsheet_select(chartsheet : Chartsheet*) : Void
    fun chartsheet_hide(chartsheet : Chartsheet*) : Void
    fun chartsheet_set_first_sheet(chartsheet : Chartsheet*) : Void
    fun chartsheet_set_tab_color(chartsheet : Chartsheet*, color : Color) : Void
    fun chartsheet_protect(chartsheet : Chartsheet*, password : Str, options : Protection*) : Void
    fun chartsheet_set_zoom(chartsheet : Chartsheet*, scale : UInt16) : Void
    fun chartsheet_set_landscape(chartsheet : Chartsheet*) : Void
    fun chartsheet_set_portrait(chartsheet : Chartsheet*) : Void
    fun chartsheet_set_paper(chartsheet : Chartsheet*, paper_type : UInt8) : Void
    fun chartsheet_set_margins(chartsheet : Chartsheet*, left : LibC::Double, right : LibC::Double, top : LibC::Double, bottom : LibC::Double) : Void
    fun chartsheet_set_header(chartsheet : Chartsheet*, string : Str) : LXWError
    fun chartsheet_set_footer(chartsheet : Chartsheet*, string : Str) : LXWError
    fun chartsheet_set_header_opt(chartsheet : Chartsheet*, string : Str, options : HeaderFooterOptions*) : LXWError
    fun chartsheet_set_footer_opt(chartsheet : Chartsheet*, string : Str, options : HeaderFooterOptions*) : LXWError

    # utility.h
    fun lxw_version : Str
    fun lxw_strerror(error_num : LXWError) : Str
  end
end