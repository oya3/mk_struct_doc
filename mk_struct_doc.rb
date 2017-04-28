# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

require 'optparse'
require 'outputer_axlsx'
require 'collector/struct_collector'
require 'collector/enum_collector'
require 'outputter/axlsx_outputter'

require 'pry'

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'

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

# 各定義取得
struct_collector = StructCollector.new body # 構造体
enum_collector = EnumCollector.new body # enum

# 各定義書き出し
axlsx_outputter = AXLSXOutputter.new 
axlsx_outputter.add_struct struct_collector.documents, struct_collector.contents
axlsx_outputter.add_enum enum_collector.documents, enum_collector.contents
axlsx_outputter.save argv[1]
#create_xls_for_axlsx(struct_collector.documents, struct_collector.contents, argv[1] + ".xlsx")
exit

# doxygen commnet 取得
documents = Hash.new
body.gsub(/^\/\*\*[\t \*\n]+?@struct\s+(.+?)\s*\n(.+?)\*\/$/m) do
  title = $1
  document = $2
  params = Hash.new
  document.gsub(/@(.+?)\s+([^@]+)/m) do
    key = $1
    value = $2
    value.gsub!(/[ \t]+\*[ \t]+/,'')
    value = value.strip
    params[key] = value
  end
  documents[title] = params
end

# typedef struct 構造抽出
typedefs = Hash.new # 全構造体保持用
multi_typedef_names = Hash.new # 多重定義名置き換え用
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

create_xls_for_axlsx(documents, typedefs, argv[1] + ".xlsx")
#create_new_header(body, multi_typedef_names, argv[0]+".new") # 新しい構造定義ファイル

puts 'complete.'

