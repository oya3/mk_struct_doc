# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'

class BaseCollector
  attr_reader :documents, :contents
  def initialize(source)
    @body = source.clone
    get_documents
    get_contents
  end
end

