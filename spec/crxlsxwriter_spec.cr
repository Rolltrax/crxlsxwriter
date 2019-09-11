require "./spec_helper"



module CrXLSXWriter
  Workbook = LibXLSXWriter.workbook_new("test.xlsx")
  Worksheet = LibXLSXWriter.workbook_add_worksheet(Workbook, "Test")

  @@data_cursor = 100
  
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

  it "writes formula to cell" do
    row, col = get_row_col()
    LibXLSXWriter.worksheet_write_formula(Worksheet, row, col, "=SUM(1, 2)", nil)
  end

  it "writes datetime to cell" do
    row, col = get_row_col()
    datetime = LibXLSXWriter::Datetime.new
    datetime.year = 2001
    datetime.month = 1
    datetime.day = 16
    LibXLSXWriter.worksheet_write_datetime(Worksheet, row, col, pointerof(datetime), nil)
  end

  it "writes url to cell" do
    row, col = get_row_col()
    LibXLSXWriter.worksheet_write_url(Worksheet, row, col, "https://rolltrax.com", nil)
  end

  it "writes blank to cell" do
    row, col = get_row_col()
    LibXLSXWriter.worksheet_write_blank(Worksheet, row, col, nil)
  end

  it "writes formula number to cell" do 
    row, col = get_row_col()
    LibXLSXWriter.worksheet_write_formula_num(Worksheet, row, col, "=1 + 2", nil, 3)
  end

  it "writes rich string to cell" do
    row, col = get_row_col()
    rich1 = LibXLSXWriter::RichString.new
    rich1.string = "Is bold!"
    format1 = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_bold(format1)
    rich1.format = format1
    rich2 = LibXLSXWriter::RichString.new
    rich2.string = "Is Under!"
    format2 = LibXLSXWriter.workbook_add_format(Workbook)
    LibXLSXWriter.format_set_underline(format2, LibXLSXWriter::UnderlineStyle::LXW_UNDERLINE_SINGLE)
    rich2.format = format2
    riches = StaticArray(LibXLSXWriter::RichString*, 16).new(Pointer(LibXLSXWriter::RichString).null)
    riches[0] = pointerof(rich1)
    riches[1] = pointerof(rich1)
    riches[2] = pointerof(rich2)
    riches[3] = pointerof(rich2)
    riches[4] = pointerof(rich2)
    riches[5] = pointerof(rich2)
    riches[6] = pointerof(rich1)
    riches[7] = pointerof(rich1)
    riches[8] = pointerof(rich1)
    LibXLSXWriter.worksheet_write_rich_string(Worksheet, row, col, riches, nil)
  end

  it "writes an image to cell with options" do
    row, col = get_row_col()
    options = LibXLSXWriter::ImageOptions.new
    options.x_scale = 0.1
    options.y_scale = 0.1
    LibXLSXWriter.worksheet_insert_image_opt(Worksheet, row, col, "logo-new.png", pointerof(options))
  end

  it "creates and inserts a chart (bar)" do 
    row, col = get_row_col()
    chart = LibXLSXWriter.workbook_add_chart(Workbook, LibXLSXWriter::ChartType::LXW_CHART_COLUMN)
    LibXLSXWriter.worksheet_write_number(Worksheet, @@data_cursor.to_u32, 0.to_u16, 5, nil)
    LibXLSXWriter.worksheet_write_number(Worksheet, @@data_cursor.to_u32, 1.to_u16, 7, nil)
    LibXLSXWriter.worksheet_write_number(Worksheet, @@data_cursor.to_u32, 2.to_u16, 1, nil)
    series = LibXLSXWriter.chart_add_series(chart, nil, nil)
    LibXLSXWriter.chart_series_set_values(series, "Test", @@data_cursor, 0, @@data_cursor, 2)
    LibXLSXWriter.worksheet_insert_chart(Worksheet, row, col, chart)
  end


  LibXLSXWriter.workbook_close(Workbook)
end
