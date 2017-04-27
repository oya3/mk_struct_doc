# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

require 'axlsx'

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'

def create_xls_for_axlsx(documents, typedefs, file_name)
  book = Axlsx::Package.new
  styles = book.workbook.styles
  headline_style = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                     :sz => 20,
                                     :b => true)
  
  headline_item_style = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                          :sz => 10,
                                          :alignment => { :horizontal => :left,
                                                          :vertical => :center,
                                                          :wrap_text => true
                                                        } )
  
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

  # 全データをなめてエクセル表示位置を確認する
  postions = Hash.new
  # セルは１始まり(1) + (構造体一覧(1) + 見出し(1)) + 構造体数(n) + 改行(1)
  offset = 1 + 2 + typedefs.length + 1
  typedefs.each do |name,values|
    postions[name] = offset
    # 前回 + (構造体名(1)+brief(1)+details(1)+見出し(1)) + メンバ数(n) + 改行(1)
    offset = offset + 4 + values.length + 1 
  end
  
  
  # 構造体一覧シート
  sheet = book.workbook.add_worksheet( :name => '構造体一覧')
  sheet.add_row(['構造体一覧'], :style => headline_style)
  sheet.add_row(['No.', '構造体名', '備考'], :style => title_style)
  index = 0
  typedefs.each do |name,values|
    index += 1
    # name_link = "=HYPERLINK(\"\##{name}!A1\",\"#{name}\")"
    name_link = "=HYPERLINK(\"\#A#{postions[name]}\",\"#{name}\")"
    doc = ''
    if documents.has_key? name
      doc = documents[name]['brief'] || ''
    end
    sheet.add_row([ index, name_link, doc], :style => item_style)
  end
  sheet.column_widths 5, 50, 50, 50 # カラム指定
  sheet.add_row(['']) # 改行

  # 構造体's
  typedefs.each do |name,values|
    # puts "add worksheet #{name}"
    # sheet = book.workbook.add_worksheet( :name => name )
    sheet.add_row([name + "構造体"], :style => headline_style)
    brief = ''
    details = ''
    if documents.has_key? name
      brief = documents[name]['brief'] || ''
      details = documents[name]['details'] || ''
    end
    sheet.add_row(['概要：',brief], :style => headline_item_style)
    sheet.add_row(['詳細：',details], :style => headline_item_style)
    sheet.add_row(['No.', 'メンバ型', 'メンバ変数名', '備考'], :style => title_style)
    values.each.with_index(1) do |params, _index|
      type_link = params[:type]
      if( typedefs.has_key? params[:type] )
        type_link = "=HYPERLINK(\"\#A#{postions[params[:type]]}\",\"#{params[:type]}\")"
      end
      sheet.add_row([_index, type_link, params[:name], params[:comment]], :style => item_style)
    end
    # column 幅
    sheet.column_widths 5, 50, 50, 50 # カラム指定
    sheet.add_row(['']) # 改行
  end
  book.serialize(file_name) # 書き出し
end

