module CrXLSXWriter
  @[Link("xlsxwriter")]
  lib LibXLSXWriter
    type Workbook = Void*
    type Worksheet = Void*
    type Format = Void*
    type Chart = Void*
    type Series = Void*

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

    enum ChartType
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

    enum ChartLegendPosition
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

    enum ChartLineDashType
      LXW_CHART_LINE_DASH_SOLID
      LXW_CHART_LINE_DASH_ROUND_DOT
      LXW_CHART_LINE_DASH_SQUARE_DOT
      LXW_CHART_LINE_DASH_DASH
      LXW_CHART_LINE_DASH_DASH_DOT
      LXW_CHART_LINE_DASH_LONG_DASH
      LXW_CHART_LINE_DASH_LONG_DASH_DOT
      LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT
    end

    enum ChartMarkerType
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

    enum ChartPatternType
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

    enum ChartLabelPosition
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

    enum ChartLabelSeperator
      LXW_CHART_LABEL_SEPARATOR_COMMA 	
      LXW_CHART_LABEL_SEPARATOR_SEMICOLON 	
      LXW_CHART_LABEL_SEPARATOR_PERIOD 	
      LXW_CHART_LABEL_SEPARATOR_NEWLINE 	
      LXW_CHART_LABEL_SEPARATOR_SPACE
    end

    enum ChartAxisType
      LXW_CHART_AXIS_TYPE_X
      LXW_CHART_AXIS_TYPE_Y
    end

    enum ChartAxisTickPosition
      LXW_CHART_AXIS_POSITION_ON_TICK 	
      LXW_CHART_AXIS_POSITION_BETWEEN
    end

    enum ChartAxisLabelPosition
      LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO 	
      LXW_CHART_AXIS_LABEL_POSITION_HIGH 	
      LXW_CHART_AXIS_LABEL_POSITION_LOW 	
      LXW_CHART_AXIS_LABEL_POSITION_NONE
    end

    enum ChartAxisLabelAlignment
      LXW_CHART_AXIS_LABEL_ALIGN_CENTER 	
      LXW_CHART_AXIS_LABEL_ALIGN_LEFT 	
      LXW_CHART_AXIS_LABEL_ALIGN_RIGHT
    end

    enum ChartAxisDisplayUnit
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

    enum ChartAxisTickMark
      LXW_CHART_AXIS_TICK_MARK_DEFAULT 	
      LXW_CHART_AXIS_TICK_MARK_NONE 	
      LXW_CHART_AXIS_TICK_MARK_INSIDE 	
      LXW_CHART_AXIS_TICK_MARK_OUTSIDE 	
      LXW_CHART_AXIS_TICK_MARK_CROSSING
    end
    
    enum ChartBlank
      LXW_CHART_BLANKS_AS_GAP 	
      LXW_CHART_BLANKS_AS_ZERO 	
      LXW_CHART_BLANKS_AS_CONNECTED 
    end

    enum ChartErrorBarType
      LXW_CHART_ERROR_BAR_TYPE_STD_ERROR 	
      LXW_CHART_ERROR_BAR_TYPE_FIXED 	
      LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE 	
      LXW_CHART_ERROR_BAR_TYPE_STD_DEV 
    end

    enum ChartErrorBarAxis
      LXW_CHART_ERROR_BAR_AXIS_X 	
      LXW_CHART_ERROR_BAR_AXIS_Y
    end
    
    enum ChartErrorBarCap
      LXW_CHART_ERROR_BAR_END_CAP 	
      LXW_CHART_ERROR_BAR_NO_CAP
    end

    enum ChartTrendlineType
      LXW_CHART_TRENDLINE_TYPE_LINEAR 	
      LXW_CHART_TRENDLINE_TYPE_LOG 	
      LXW_CHART_TRENDLINE_TYPE_POLY 	
      LXW_CHART_TRENDLINE_TYPE_POWER 	
      LXW_CHART_TRENDLINE_TYPE_EXP 	
      LXW_CHART_TRENDLINE_TYPE_AVERAGE
    end

    enum ChartErrorBarDirection
      LXW_CHART_ERROR_BAR_DIR_BOTH 	
      LXW_CHART_ERROR_BAR_DIR_PLUS 	
      LXW_CHART_ERROR_BAR_DIR_MINUS 
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
      string : UInt8*
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

    fun workbook_new(path : UInt8*) : Workbook*
    fun workbook_add_worksheet(workbook : Workbook*, sheetname : UInt8*) : Worksheet*
    fun workbook_close(workbook : Workbook*) : LXWError

    fun workbook_add_format(workbook : Workbook*) : Format*
    fun workbook_add_chart(workbook : Workbook*, chart_type : ChartType) : Chart*

    fun worksheet_write_string(worksheet : Worksheet*, row : Row, col : Col, string : UInt8*, format : Format*) : LXWError
    fun worksheet_write_number(worksheet : Worksheet*, row : Row, col : Col, value : LibC::Double, format : Format*) : LXWError
    fun worksheet_write_formula(worksheet : Worksheet*, row : Row, col : Col, value : UInt8*, format : Format*) : LXWError
    # workbook_write_array_formula
    fun worksheet_write_datetime(worksheet : Worksheet*, row : Row, col : Col, datetime : Datetime*, format : Format*) : LXWError
    fun worksheet_write_url(worksheet : Worksheet*, row : Row, col : Col, url : UInt8*, format : Format*) : LXWError
    # worksheet_write_boolean
    fun worksheet_write_blank(worksheet : Worksheet*, row : Row, col : Col, format : Format*) : LXWError
    fun worksheet_write_formula_num(worksheet : Worksheet*, row : Row, col : Col, formula : UInt8*, format : Format*, result : LibC::Double) : LXWError
    fun worksheet_write_rich_string(worksheet : Worksheet*, row : Row, col : Col, rich_string : StaticArray(RichString*, 16), format : Format*) : LXWError

    fun worksheet_set_row(worksheet : Worksheet*, row : Row, height : LibC::Double, format : Format*) : LXWError
    fun worksheet_set_row_opt(worksheet : Worksheet*, row : Row, height : LibC::Double, format : Format*, options : RowColOptions*) : LXWError
    fun worksheet_set_column(worksheet : Worksheet*, first_col : Col, last_col : Col, width : LibC::Double, format : Format*) : LXWError
    fun worksheet_set_column_opt(worksheet : Worksheet*, first_col : Col, last_col : Col, width : LibC::Double, format : Format*, options : RowColOptions*) : LXWError

    fun worksheet_insert_image(worksheet : Worksheet*, row : Row, col : Col, filename : UInt8*) : LXWError
    fun worksheet_insert_image_opt(worksheet : Worksheet*, row : Row, col : Col, filename : UInt8*, options : ImageOptions*) : LXWError
    fun worksheet_insert_chart(worksheet : Worksheet*, row : Row, col : Col, chart : Chart*)

    fun format_set_bold(format : Format*) : Void
    fun format_set_font_color(format : Format*, color : Color)
    fun format_set_font_name(format : Format*, font_name : UInt8*)
    fun format_set_font_size(format : Format*, font_size : LibC::Double)
    fun format_set_italic(format : Format*)
    fun format_set_underline(format : Format*, style : UnderlineStyle)
    fun format_set_font_strikeout(format : Format*)
    fun format_set_font_script(format : Format*, script : FontScript)
    fun format_set_num_format(format : Format*, num_format : UInt8*)
    fun format_set_unlocked(format : Format*)
    fun format_set_hidden(format : Format*)
    fun format_set_align(format : Format*, align : Alignment)
    fun format_set_text_wrap(format : Format*)
    fun format_set_rotation(format : Format*, angle : Int16)
    fun format_set_indent(format : Format*, level : UInt8)
    fun format_set_pattern(format : Format*, pattern : Pattern)
    fun format_set_bg_color(format : Format*, color : Color)
    fun format_set_fg_color(format : Format*, color : Color)
    fun format_set_fg_color(format : Format*, color : Color)
    fun format_set_border(format : Format*, style : Border)
    fun format_set_bottom(format : Format*, style : Border)
    fun format_set_top(format : Format*, style : Border)
    fun format_set_left(format : Format*, style : Border)
    fun format_set_right(format : Format*, style : Border)
    fun format_set_border_color(format : Format*, color : Color)
    fun format_set_bottom_color(format : Format*, color : Color)
    fun format_set_top_color(format : Format*, color : Color)
    fun format_set_left_color(format : Format*, color : Color)
    fun format_set_right_color(format : Format*, color : Color)

    fun chart_add_series(chart : Chart*, categories : UInt8*, values : UInt8*) : Series*
    fun chart_series_set_values(chart : Series*, sheetname : UInt8*, first_row : Row, first_col : Col, last_row : Row, last_col : Col)
  end
end
