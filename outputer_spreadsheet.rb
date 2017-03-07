# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

require 'spreadsheet'

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'
Spreadsheet.client_encoding = 'utf-8' # spreadsheet

# Spreadsheet cellフォーマット
class SpreadsheetFontFormat < Spreadsheet::Format
  def initialize(fontName,fontSize,hash={})
    super()
    hash.each do | (key,value) |
      self.method("#{key}=").call(value) # メソッドの場合
    end
    self.font = Spreadsheet::Font.new(fontName, size: fontSize)
  end
end

# 構造体出力
def make_sheet_struct(sheet, typedefs, name, values)
  # haedline format
  headline_format = SpreadsheetFontFormat.new( 'ＭＳ Ｐゴシック', 20,
                                               { :left => :thin,
                                                 :right => :thin,
                                                 :top => :thin,
                                                 :bottom => :thin,
                                                 :text_wrap => false } )
  headline_format.font.weight = :bold
  # title format
  title_format = SpreadsheetFontFormat.new( 'ＭＳ Ｐゴシック', 10,
                                            { :left=>:thin,
                                              :right=>:thin,
                                              :top=>:thin,
                                              :bottom=>:thin,
                                              :pattern => 1,
                                              :pattern_fg_color => :gray,
                                              :text_wrap => true } ) # :text_wrap = true でないと改行できない
  title_format.font.weight = :bold
  
  # sheet[y,x]
  # 構造体名
  sheet[0, 0] = name + " 構造体"
  sheet.row(0).set_format(0, headline_format )
  sheet.row(0).height = 40
  
  # 項目タイトル
  y = 1
  sheet[y, 0] = 'No.'
  sheet[y, 1] = 'メンバ型'
  sheet[y, 2] = 'メンバ変数名'
  sheet[y, 3] = '備考'
  
  # フォーマット
  sheet.row(y).set_format(0, title_format )
  sheet.row(y).set_format(1, title_format )
  sheet.row(y).set_format(2, title_format )
  sheet.row(y).set_format(3, title_format )
  
  # row 幅
  sheet.row(y).height = 20
  # column 幅
  sheet.column(0).width = 5
  sheet.column(1).width = 50
  sheet.column(2).width = 50
  sheet.column(3).width = 50
  
  # item format
  item_format = SpreadsheetFontFormat.new( 'ＭＳ Ｐゴシック', 10,
                                           { :left => :thin,
                                             :right => :thin,
                                             :top => :thin,
                                             :bottom => :thin,
                                             :text_wrap => true } )
  
  values.each.with_index(1) do |params,index|
    y += 1
    sheet[y, 0] = index
    if( typedefs.has_key? params[:type] )
      sheet[y, 1] = "=HYPERLINK(\"\##{params[:type]}!A1\",\"#{params[:type]}\")"
    else
      sheet[y, 1] = params[:type]
    end
    sheet[y, 2] = params[:name]
    sheet[y, 3] = params[:comment]
    # フォーマット
    sheet.row(y).set_format(0, item_format )
    sheet.row(y).set_format(1, item_format )
    sheet.row(y).set_format(2, item_format )
    sheet.row(y).set_format(3, item_format )
    # row 幅
    sheet.row(y).height = 20
  end
end

# 構造体一覧出力
def make_sheet_struct_list(sheet,typedefs)
  # haedline format
  headline_format = SpreadsheetFontFormat.new( 'ＭＳ Ｐゴシック', 20,
                                               { :left => :thin,
                                                 :right => :thin,
                                                 :top => :thin,
                                                 :bottom => :thin,
                                                 :text_wrap => false } )
  headline_format.font.weight = :bold
  # title format
  title_format = SpreadsheetFontFormat.new( 'ＭＳ Ｐゴシック', 10,
                                            { :left=> :thin,
                                              :right=> :thin,
                                              :top=> :thin,
                                              :bottom=> :thin,
                                              :pattern => 1,
                                              :pattern_fg_color => :gray,
                                              :text_wrap => true } ) # :text_wrap = true でないと改行できない
  title_format.font.weight = :bold
  
  # sheet[y,x]
  # 構造体名
  sheet[0, 0] = "構造体一覧"
  sheet.row(0).set_format(0, headline_format )
  sheet.row(0).height = 40
  
  # 項目タイトル
  y = 1
  sheet[y, 0] = 'No.'
  sheet[y, 1] = '構造体名'
  sheet[y, 2] = '備考'
  
  # フォーマット
  sheet.row(y).set_format(0, title_format )
  sheet.row(y).set_format(1, title_format )
  sheet.row(y).set_format(2, title_format )
  
  # column 幅
  sheet.column(0).width = 5
  sheet.column(1).width = 50
  sheet.column(2).width = 50
  # row 幅
  sheet.row(y).height = 20

  # item format
  item_format = SpreadsheetFontFormat.new( 'ＭＳ Ｐゴシック', 10,
                                           { :left => :thin,
                                             :right => :thin,
                                             :top => :thin,
                                             :bottom => :thin,
                                             :text_wrap => true } )
  index = 0
  typedefs.each do |name,values|
    index += 1
    y += 1
    sheet[y, 0] = index
    sheet[y, 1] = "=HYPERLINK(\"\##{name}!A1\",\"#{name}\")"
    sheet[y, 2] = ''
    # フォーマット
    sheet.row(y).set_format(0, item_format )
    sheet.row(y).set_format(1, item_format )
    sheet.row(y).set_format(2, item_format )
    # row 幅
    sheet.row(y).height = 20
  end
end

def create_xls_for_spreadsheet(typedefs, file_name)
  begin
    book = Spreadsheet::Workbook.new
    # 見出
    sheet = book.create_worksheet( :name => "構造体一覧" )
    sheet.default_format = SpreadsheetFontFormat.new('ＭＳ Ｐゴシック', 10)
    make_sheet_struct_list(sheet,typedefs)
    # 関数
    typedefs.each do |name,values|
      sheet = book.create_worksheet( :name => name )
      sheet.default_format = SpreadsheetFontFormat.new('ＭＳ Ｐゴシック', 10) # sheet に設定する場合
      make_sheet_struct(sheet, typedefs, name, values)
    end
    book.write file_name
  rescue => e
    puts "#{__method__} error:" + e.message
    exit
  end
end
