# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))
require 'base_collector'

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'

class EnumCollector < BaseCollector
  def initialize(source)
    super source
  end

  # doxygen commnet 取得
  def get_documents
    @documents = Hash.new
    @body.gsub(/^\/\*\*[\t \*\n]+?@enum[ \t]+(.+?)[ \t]*\n(.+?)\*\/$/m) do
      title = $1
      document = $2
      params = Hash.new
      document.gsub(/@(.+?)[ \t]+([^@]+)/m) do
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
    # typedef enum 構造抽出
    @contents = Hash.new # 全構造体保持用
    @body.gsub(/^[ \t]*typedef[ \t]+enum[ \t]*.*?\{(.+?)\}[ \t]*(.+?);/m) do |match|
      # puts "get defines  #{$1} #{$2}"
      content = $1.clone
      content_names = $2.strip.split(',')
      values = Array.new

      content.gsub(/^[ \t]*([a-zA-Z_][a-zA-Z0-9_]+)(.*)\n/) do |line|
        name = $1
        value = $2
        params = Hash.new
        params[:type] = name.strip
        params[:value] = value[/^[ \t]*=[ \t]*(.+?)[ \t]*\,/,1] || value[/^[ \t]*=[ \t]*(.+?)[ \t]*(\/\*|\/\/)/,1]
        params[:comment] = value[/^.+\/\/[ \t]*(.+?)[ \t]*\n/, 1] || value[/^.+\/\*[ \t]*(.+?)[ \t]*\*\//,1]
        params[:names] = content_names;
        values << params;
      end
      @contents[content_names[0]] = values;
    end
  end
end

