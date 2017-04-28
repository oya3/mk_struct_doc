# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

require 'base_collector'

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'

class StructCollector < BaseCollector
  attr_reader :multi_typedef_names # TODO: 複数名定義保持用エリア
  def initialize(source)
    super source
  end

  # doxygen commnet 取得
  def get_documents
    @documents = Hash.new
    @body.gsub(/^\/\*\*[\t \*\n]+?@struct\s+(.+?)\s*\n(.+?)\*\/$/m) do
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
      @documents[title] = params
    end
  end

  def get_contents
    # typedef struct 構造抽出
    @contents = Hash.new # 全構造体保持用
    @multi_typedef_names = Hash.new # 多重定義名置き換え用
    @body.gsub(/^\s*typedef\s+struct\s*.*?\{(.+?)\}\s*(.+?);/m) do |match|
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
      @contents[one_typedef_names[0]] = values;
      
      # names = one_typedef_names.strip.split(",")
      if(one_typedef_names.size != 1)
        1.upto(one_typedef_names.size - 1) do |index|
          @multi_typedef_names[one_typedef_names[index].strip] = one_typedef_names[0].strip
        end
      end
    end
  end
end

