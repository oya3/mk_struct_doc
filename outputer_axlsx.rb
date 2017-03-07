# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

require 'axlsx'

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'

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

