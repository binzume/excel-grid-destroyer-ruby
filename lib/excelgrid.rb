# encoding: utf-8

require 'rexml/document'
require 'zip'

module ExcelGrid
  class Book
    attr_reader :zip_file, :ss, :theme_color
    def initialize(file_name)
      @zip_file = Zip::File.open(file_name)
      ent = @zip_file.get_entry('xl/sharedStrings.xml')
      strings = REXML::Document.new(ent.get_input_stream.read)
      # strings = REXML::Document.new(open("tmp/#{name}/xl/sharedStrings.xml"))
      @ss = []
      strings.elements.each('sst/si') {|str|
        if str.elements['t']
          @ss << str.elements['t'].text
        else
          @ss << str.elements.each('r'){|e| e.text }.join.gsub(/<[^>]+>/,'')
        end
      }
      read_theme
    end

    def read_theme
      ent = @zip_file.get_entry('xl/theme/theme1.xml')
      theme_xml = REXML::Document.new(ent.get_input_stream.read)

      @theme_color = {}
      theme_xml.elements.each('a:theme/a:themeElements/a:clrScheme/*') {|clr|
        rgb = clr.elements['a:srgbClr'] && clr.elements['a:srgbClr'].attributes['val']
        @theme_color[clr.name] = rgb || (clr.elements['a:sysClr'] && clr.elements['a:sysClr'].attributes['lastClr'])
      }
    end

    def sheet(sheet, scale)
      Sheet.new(self, sheet, scale)
    end
  end

  class Sheet
    attr_reader :cells, :drawings, :default_width, :default_height
    attr_accessor :scale

    def initialize(book, sheet, scale = 1.0)
      @book = book
      @zip_file = book.zip_file
      @scale = scale
      @sheet = sheet

      load_cell
      load_drawing
    end

    def col(cr)
      cr.bytes.inject(0) {|v,c|
        break v if c < 65
        (v) * 26 + c - 64
      }
    end

    def row(cr)
      cr.sub(/^[A-Z]+/,'').to_i
    end

    def range(crcr)
     crcr.split(':').map {|cr|
       [col(cr),row(cr)]
     }
    end

    def emu2pt(emu)
      emu / 12700.0
    end

    def load_cell
      ent = @zip_file.get_entry('xl/worksheets/' + @sheet + '.xml')
      sheet = REXML::Document.new(ent.get_input_stream.read)

      @default_width = (sheet.elements['worksheet/sheetFormatPr'].attributes['defaultColWidth'].to_f * 7 + 4) * 72 / 96
      @default_height = sheet.elements['worksheet/sheetFormatPr'].attributes['defaultRowHeight'].to_f
      @row_height = []

      cells = []
      sheet.elements.each('worksheet/sheetData/row') {|row|
        r = row.attributes['r'].to_i
        @row_height[r] = (row.attributes['ht'] || @default_height).to_f
        row.elements.each('c') {|cell|
          v = if cell.attributes['t']== 's'
            @book.ss[cell.elements['v'].text.to_i]
          else
            cell.elements['v']
          end
          c = col(cell.attributes['r'])
          cells[r] = [] unless cells[r]
          cells[r][c] = {:v => v, :col => c, :row=>r, :id=>cell.attributes['r']}
        }
      }

      sheet.elements.each('worksheet/mergeCells/mergeCell') {|merge|
       r = range(merge.attributes['ref'])
       #p r, merge.attributes['ref']
       if cells[r[0][1]] && cells[r[0][1]][r[0][0]]
        cells[r[0][1]][r[0][0]][:w] = r[1][0] - r[0][0] + 1;
        cells[r[0][1]][r[0][0]][:h] = r[1][1] - r[0][1] + 1;
       end
      }
      @cells = cells
      @row_pos_y =  @row_height.each_with_object([0]){|h,a|
        a << a.last + (h || @default_height)
      }
    end

    def col_x(c)
        c * @default_width
    end

    def row_y(r)
        @row_pos_y[r] || r * @default_height
    end

    def load_drawing
      ent = @zip_file.get_entry('xl/drawings/drawing1.xml')
      drawing = REXML::Document.new(ent.get_input_stream.read)
      scale = @scale

      drawings = []
      drawing.elements.each('xdr:wsDr/xdr:twoCellAnchor') {|cell_anchor|
        if cell_anchor.elements['xdr:sp/xdr:spPr/a:prstGeom']
          from_col = cell_anchor.elements['xdr:from/xdr:col'].text.to_i
          from_row = cell_anchor.elements['xdr:from/xdr:row'].text.to_i
          from_col_off = cell_anchor.elements['xdr:from/xdr:colOff'].text.to_i
          from_row_off = cell_anchor.elements['xdr:from/xdr:rowOff'].text.to_i
          to_col = cell_anchor.elements['xdr:to/xdr:col'].text.to_i
          to_row = cell_anchor.elements['xdr:to/xdr:row'].text.to_i
          to_col_off = cell_anchor.elements['xdr:to/xdr:colOff'].text.to_i
          to_row_off = cell_anchor.elements['xdr:to/xdr:rowOff'].text.to_i
          fillcolor = nil
          theme_color = cell_anchor.elements['xdr:sp/xdr:spPr/a:solidFill/a:schemeClr']
          if theme_color
            name = theme_color.attributes['val']
            fillcolor = name && @book.theme_color[name]
          end
          geom = cell_anchor.elements['xdr:sp/xdr:spPr/a:prstGeom'].attributes['prst']
          s = cell_anchor.elements.each('.//a:t'){|e| e}.map{|e| e.text}.join
          #puts geom + ":" + s
          x = ((from_col * default_width + emu2pt(from_col_off)) * scale).round(2)
          y = ((row_y(from_row) + emu2pt(from_row_off)) * scale).round(2)
          w = ((to_col * default_width + emu2pt(to_col_off)) * scale).round(2) - x
          h = ((row_y(to_row) + emu2pt(to_row_off)) * scale).round(2) - y
          style = "top:#{y}pt;left:#{x}pt;width:#{w}pt;height:#{h}pt;"
          if fillcolor
            style += "background-color: rgba(#{fillcolor[0,2].hex},#{fillcolor[2,2].hex},#{fillcolor[4,2].hex},0.6);"
          end
          if geom == 'roundRect'
            style += 'border-radius:10pt;';
          elsif geom == 'flowChartAlternateProcess'
            style += "border-radius:4pt;";
          elsif geom == 'ellipse'
            style += "border-radius:#{w}pt;";
          end

          #p default_width
          #p  emu2pt(from_col_off)
          #puts s
          drawings << {type: 'two_cell', from_col: from_col, from_row: from_row, height: h, width: w, style: style, text: s}
        end
      }

      # TODO:
      #drawing.elements.each('xdr:wsDr/xdr:oneCellAnchor') {|cell_anchor|
      #  puts "oneCellAnchor:" + cell_anchor.elements['//a:prstGeom'].attributes['prst']
      #   drawings << {type: 'one_cell', col: col, row: row, style: style, text: s}
      #}

      @drawings = drawings
      drawing
    end

    def to_html(id, css_class)
      scale = @scale
      #default_width = @default_width
      #default_height = @default_height

      html = ""
      drawings.each {|drawing|
          html += "<div class='#{css_class}' id='#{id}_#{drawing[:from_col]}.#{drawing[:from_row]}' style='#{drawing[:style]}'><span style='height:#{drawing[:height]}pt'>#{drawing[:text]}</span></div>\n"
      }

      cells.each{|row|
       row.each{|cell|
         if cell
           w = ((cell[:w] || 1) * default_width * scale).round(2)
           h = (row_y(cell[:h] || 1) * scale).round(2)
           x = ((cell[:col]-1) * default_width * scale).round(2)
           y = (row_y(cell[:row]-1) * scale).round(2)
           if (cell[:v] || cell[:h])
             name = cell[:v].to_s.sub(/[ 　]+/,' ').sub(/^\d+(?=[^-\d])/,'').strip.sub(/\n/,'<br />')
             html += "<div class='#{css_class}' id='#{id}_#{cell[:id]}' style='top:#{y}pt;left:#{x}pt;width:#{w}pt;height:#{h}pt'><span style='height:#{h}pt'>#{name}</span></div>\n"
            end
         end
       } if row
      }
      html
    end

  end

end

