module CrXLSXWriter
  @[Link("xlsxwriter")]
  lib LibXLSXWriter
    type Workbook = Void*
    type Worksheet = Void*
    type Format = Void*

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
      LXW_UNDERLINE_SINGLE = 1
      LXW_UNDERLINE_DOUBLE = 2
      LXW_UNDERLINE_SINGLE_ACCOUNTING = 3
      LXW_UNDERLINE_DOUBLE_ACCOUNTING = 4
    end

    enum FontScript : Int8
      LXW_FONT_SUPERSCRIPT = 1
      LXW_FONT_SUBSCRIPT = 2
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

    fun workbook_new(path : UInt8*) : Workbook*
    fun workbook_add_worksheet(workbook : Workbook*, sheetname : UInt8*) : Worksheet*
    fun workbook_close(workbook : Workbook*) : LXWError
    fun workbook_add_format(workbook : Workbook*) : Format*

    fun worksheet_write_string(worksheet : Worksheet*, row : Row, col : Col, string : UInt8*, format : Format*)
    fun worksheet_write_number(worksheet : Worksheet*, row : Row, col : Col, value : LibC::Double, format : Format*)

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
  end
end
