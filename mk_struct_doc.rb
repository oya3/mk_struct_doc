# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

require "optparse"
require 'spreadsheet'
require 'axlsx'

require "pry"

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

# メンバ変数型名を新定義型名に変更
def create_new_header(body, multi_typedef_names, file_name)
  new_body = body.clone
  multi_typedef_names.each do |key, value|
    new_body.gsub!(/^\s+#{key}\s+.+?;/) do |line|
      line.gsub!(/#{key}/,value)
    end
  end
  File.open(file_name, "w:cp932:utf-8") do |file|
    file.write new_body
  end
end

# 
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

def create_xls_for_axlsx(typedefs, file_name)
  book = Axlsx::Package.new
  styles = book.workbook.styles
  headline_style = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                     :sz => 20,
                                     :b => true)
  title_style = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                  :sz => 10,
                                  :bg_color => 'c0c0c0',
                                  :b => true,
                                  :border => { :style => :thin, :color => "00" },
                                  :alignment => { :horizontal => :center,
                                                  :vertical => :center ,
                                                  :wrap_text => true})
  item_style = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                 :sz => 10,
                                 :border => { :style => :thin, :color => "00" },
                                 :alignment => { :horizontal => :left,
                                                 :vertical => :center ,
                                                 :wrap_text => true})

  # 構造体一覧シート
  sheet = book.workbook.add_worksheet( :name => '構造体一覧')
  sheet.add_row(['構造体一覧'], :style => headline_style)
  sheet.add_row(['No.', '構造体名', '備考'], :style => title_style)
  index = 0
  typedefs.each do |name,values|
    index += 1
    name_link = "=HYPERLINK(\"\##{name}!A1\",\"#{name}\")"
    sheet.add_row([index,name_link,''], :style => item_style)
  end
  sheet.column_widths 5, 30, 30 # カラム指定

  # 構造体シート's
  typedefs.each do |name,values|
    sheet = book.workbook.add_worksheet( :name => name )
    sheet.add_row([name + "構造体"], :style => headline_style)
    sheet.add_row(['No.', 'メンバ型', 'メンバ変数名', '備考'], :style => title_style)
    values.each.with_index(1) do |params, _index|
      type_link = params[:type]
      if( typedefs.has_key? params[:type] )
        type_link = "=HYPERLINK(\"\##{params[:type]}!A1\",\"#{params[:type]}\")"
      end
      sheet.add_row([_index, type_link, params[:name], params[:comment]], :style => item_style)
    end
    # column 幅
    sheet.column_widths 5, 50, 50, 50 # カラム指定
  end
  book.serialize(file_name) # 書き出し
end

# main

# オプション
options = Hash.new
opt = OptionParser.new
# opt.on('-b VAL', '--byte VAL') {|val| options[:b] = val }
# opt.on('-d', '--debug') {|val| options[:debug] = val }

argv = opt.parse(ARGV)
if argv.length != 2
  abort "usage: mk_struct_doc [header<.h>] [output<.xls>]"
end

body = File.read( argv[0], encoding: 'cp932:utf-8' )

typedefs = Hash.new # 全構造体保持用
multi_typedef_names = Hash.new # 多重定義名置き換え用
# typedef struct 構造抽出
body.gsub(/^\s*typedef\s+struct\s*.*?\{(.+?)\}\s*(.+?);/m) do |match|
  # puts "get defines  #{$1} #{$2}"
  one_typedef = $1.clone
  one_typedef_names = $2.strip.split(',')
  values = Array.new
  one_typedef.gsub(/^\s*(([a-zA-Z_][a-zA-Z0-9_]*\s+){1,})([a-zA-Z_][a-zA-Z0-9_\[\]]*?)\s*;(.*?)\n/) do |line|
    params = Hash.new
    # $1 定義全て
    # $2 定義の最終のみ
    # $3 変数名
    # $4 コメント
    params[:type] = $1.strip
    params[:name] = $3.strip
    params[:comment] = $4.strip # 前後空欄削除
    params[:comment].gsub!(/\/\*\s*(.+?)\s*\*\//) { $1 } # cコメント
    params[:comment].gsub!(/\/\/\!\<\s*(.+?)$/) { $1 } # cppコメント(doxygen)
    params[:comment].gsub!(/\/\/\s*(.+?)$/) { $1 } # cppコメント
    params[:names] = one_typedef_names;
    values << params;
  end
  typedefs[one_typedef_names[0]] = values;
  
  # names = one_typedef_names.strip.split(",")
  if(one_typedef_names.size != 1)
    1.upto(one_typedef_names.size - 1) do |index|
      multi_typedef_names[one_typedef_names[index].strip] = one_typedef_names[0].strip
    end
  end
end

create_xls_for_spreadsheet(typedefs, argv[1]) # xls仕様書作成
create_xls_for_axlsx(typedefs, argv[1] + ".xlsx")
create_new_header(body, multi_typedef_names, argv[0]+".new") # 新しい構造定義ファイル

puts 'complete.'

