require "./spec_helper"



module CrXLSXWriter
  Workbook = LibXLSXWriter.workbook_new("test.xlsx")
  Worksheet = LibXLSXWriter.workbook_add_worksheet(Workbook, "Test")

  @@i = 0
  RowLength = 5
  def self.get_row_col
    row = @@i / RowLength
    col = @@i % RowLength
    @@i += 1
    return [row.to_u32, col.to_u16]
  end

  it "opens a workbook and modifies cells" do
    row, col = get_row_col()
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Example", nil)
  end

  it "writes bold string to a cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_bold(format)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Bold", format)
  end

  it "writes colored string to a cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_font_color(format, 122)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Colored", format)
  end

  it "writes different font to cell" do 
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_font_name(format, "Times New Roman")
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "TNR", format)
  end

  it "writes italic string to cell" do 
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_italic(format)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Italic", format)
  end

  it "writes underlined string to cell" do 
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_underline(format, LibXLSXWriter::UnderlineStyle::LXW_UNDERLINE_DOUBLE)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Underline", format)
  end

  it "writes struck out string to cell" do 
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_font_strikeout(format)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Strike Out!", format)
  end

  it "writes num format to cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_num_format(format, "$0.00")
    LibXLSXWriter.worksheet_write_number(Worksheet, row, col, 2.00, format)
  end

  it "writes wrapped text to cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_text_wrap(format)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Wrap this long-ass string to the cell motherfucker!", format)
  end

  it "writes rotated text to cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_rotation(format, 45.to_i16)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Rotation!", format)
  end

  it "writes indented text to cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_indent(format, 1.to_i8)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Indent", format)
  end

  it "writes superscript to cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_font_script(format, LibXLSXWriter::FontScript::LXW_FONT_SUPERSCRIPT)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "SUPER", format)
  end

  it "writes aligned text to cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_align(format, LibXLSXWriter::Alignment::LXW_ALIGN_RIGHT)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Right", format)
  end

  it "writes aligned bordered to cell" do
    row, col = get_row_col()
    format = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_border(format, LibXLSXWriter::Border::LXW_BORDER_DOUBLE)
    LibXLSXWriter.worksheet_write_string(Worksheet, row, col, "Dbl Border", format)
  end

  LibXLSXWriter.workbook_close(Workbook)
end
