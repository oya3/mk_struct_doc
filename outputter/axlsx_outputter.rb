# coding: utf-8
$LOAD_PATH.unshift(File.expand_path(File.dirname(__FILE__)))

require 'axlsx'

Encoding.default_external = 'utf-8'
Encoding.default_internal = 'utf-8'

class AXLSXOutputter
  def initialize
    @book = Axlsx::Package.new
    styles = @book.workbook.styles
    @styles = Hash.new
    @styles[:headline_style] = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                                 :sz => 20,
                                                 :b => true)
    
    @styles[:headline_item_style] = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                                      :sz => 10,
                                                      :alignment => { :horizontal => :left,
                                                                      :vertical => :center,
                                                                      :wrap_text => true
                                                                    } )
    
    @styles[:title_style] = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                              :sz => 10,
                                              :bg_color => 'c0c0c0',
                                              :b => true,
                                              :border => { :style => :thin, :color => "00" },
                                              :alignment => { :horizontal => :center,
                                                              :vertical => :center ,
                                                              :wrap_text => true})

    @styles[:item_style] = styles.add_style( :font_name => 'ＭＳ Ｐゴシック',
                                             :sz => 10,
                                             :border => { :style => :thin, :color => "00" },
                                             :alignment => { :horizontal => :left,
                                                             :vertical => :center ,
                                                             :wrap_text => true})
  end

  # 表示位置を計算
  def calc_postions(contents)
    # 全データをなめてエクセル表示位置を確認する
    postions = Hash.new
    # セルは１始まり(1) + (構造体一覧(1) + 見出し(1)) + 構造体数(n) + 改行(1)
    offset = 1 + 2 + contents.length + 1
    contents.each.with_index(1) do |(name,values), index|
      pos = Hash.new
      pos[:index] = index + 2
      pos[:start] = offset # 開始位置
      # 前回 + (構造体名(1)+brief(1)+details(1)+見出し(1)) + メンバ数(n) + (戻る(1) + 改行(1))
      offset = offset + 4 + values.length + 2
      pos[:end] = offset - 2 # 終了位置
      postions[name] = pos
    end
    return postions
  end
  
  def add_sheet_with_list(documents, contents, postions, title, item1, item2)
    sheet = @book.workbook.add_worksheet( :name => title)
    sheet.add_row([title], :style => @styles[:headline_style])
    sheet.add_row(['No.', item1, item2], :style => @styles[:title_style])
    contents.each.with_index(1) do |(name,values),index|
      cstart = postions[name][:start]
      cend = postions[name][:end]
      name_link = "=HYPERLINK(\"\#A#{cstart}:D#{cend}\",\"#{name}\")"
      doc = ''
      if documents.has_key? name
        doc = documents[name]['brief'] || ''
      end
      sheet.add_row([ index, name_link, doc], :style => @styles[:item_style])
    end
    sheet.column_widths 5, 50, 50, 50 # カラム指定
    sheet.add_row(['']) # 改行
    return sheet
  end
  
  
  # 構造体一覧を作成
  def add_struct(documents, contents)
    # 全データをなめてエクセル表示位置を確認する
    postions = calc_postions contents
    
    # 構造体一覧シート
    sheet = add_sheet_with_list documents, contents, postions, '構造体一覧', '構造体名', '備考'
    
    # 構造体's
    contents.each.with_index(1) do |(name,values),index|
      # puts "add worksheet #{name}"
      # sheet = book.workbook.add_worksheet( :name => name )
      sheet.add_row(["#{index}. " + name + "構造体"], :style => @styles[:headline_style])
      brief = ''
      details = ''
      if documents.has_key? name
        brief = documents[name]['brief'] || ''
        details = documents[name]['details'] || ''
      end
      sheet.add_row(['概要：',brief], :style => @styles[:headline_item_style])
      sheet.add_row(['詳細：',details], :style => @styles[:headline_item_style])
      sheet.add_row(['No.', 'メンバ型', 'メンバ変数名', '備考'], :style => @styles[:title_style])
      values.each.with_index(1) do |params, _index|
        type_link = params[:type]
        if( contents.has_key? params[:type] )
          cstart = postions[params[:type]][:start]
          cend = postions[params[:type]][:end]
          type_link = "=HYPERLINK(\"\#A#{cstart}:D#{cend}\",\"#{params[:type]}\")"
        end
        sheet.add_row([_index, type_link, params[:name], params[:comment]], :style => @styles[:item_style])
      end
      # column 幅
      sheet.column_widths 5, 50, 50, 50 # カラム指定
      cindex = postions[name][:index]
      list_link = "=HYPERLINK(\"\#A#{cindex}:C#{cindex}\",\"↥\")"
      sheet.add_row([list_link]) # 戻る
      sheet.add_row(['']) # 改行
    end
  end

  # enum一覧を作成
  def add_enum(documents, contents)
    # 全データをなめてエクセル表示位置を確認する
    postions = calc_postions contents
    
    # 構造体一覧シート
    sheet = add_sheet_with_list documents, contents, postions, 'enum一覧', 'enum名', '備考'
    
    # 構造体's
    contents.each.with_index(1) do |(name,values),index|
      # puts "add worksheet #{name}"
      # sheet = book.workbook.add_worksheet( :name => name )
      sheet.add_row(["#{index}. " + name], :style => @styles[:headline_style])
      brief = ''
      details = ''
      if documents.has_key? name
        brief = documents[name]['brief'] || ''
        details = documents[name]['details'] || ''
      end
      sheet.add_row(['概要：',brief], :style => @styles[:headline_item_style])
      sheet.add_row(['詳細：',details], :style => @styles[:headline_item_style])
      sheet.add_row(['No.', '定義名', '値', '備考'], :style => @styles[:title_style])
      values.each.with_index(1) do |params, _index|
        type_link = params[:type]
        if( contents.has_key? params[:type] )
          cstart = postions[params[:type]][:start]
          cend = postions[params[:type]][:end]
          type_link = "=HYPERLINK(\"\#A#{cstart}:D#{cend}\",\"#{params[:type]}\")"
        end
        sheet.add_row([_index, type_link, params[:value], params[:comment]], :style => @styles[:item_style])
      end
      # column 幅
      sheet.column_widths 5, 50, 50, 50 # カラム指定
      cindex = postions[name][:index]
      list_link = "=HYPERLINK(\"\#A#{cindex}:C#{cindex}\",\"↥\")"
      sheet.add_row([list_link]) # 戻る
      sheet.add_row(['']) # 改行
    end
  end

  def save(path)
      @book.serialize(path) # 書き出し
  end
end
  
