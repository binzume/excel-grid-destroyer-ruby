# encoding: utf-8

require 'rexml/document'
require 'zip'
require 'pathname'

module ExcelGrid
  class Book
    attr_reader :zip_file, :ss, :theme_color, :sheets
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

      bookfile = 'xl/workbook.xml'
      rel = read_rels(bookfile)

      ent = @zip_file.get_entry(bookfile)
      book_xml = REXML::Document.new(ent.get_input_stream.read)
      @sheets = []
      book_xml.elements.each('workbook/sheets/sheet') {|s|
        @sheets << {
          id: s.attributes['sheetId'],
          name: s.attributes['name'],
          state: s.attributes['state'],
          file: 'xl/' + (rel[s.attributes['r:id']] || "worksheets/sheet#{@sheets.length+1}")
        }
      }
      read_theme
    end

    def read_rels tfile
      ent = @zip_file.get_entry(tfile.sub(/([^\/]+)\Z/, "_rels/\\1.rels" ))
      return {} unless ent
      ret = {}
      doc = REXML::Document.new(ent.get_input_stream.read)
      doc.elements.each('/Relationships/Relationship') {|rel|
        ret[rel.attributes['Id']] = rel.attributes['Target']
      }
      ret
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

    def sheet(sheet, scale = 1.0)
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
      @doc = REXML::Document.new(@zip_file.get_entry(sheet).get_input_stream.read)
      @cells = load_cell()
      @drawings = load_drawing()
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
      sheet = @doc
      width = sheet.elements['worksheet/sheetFormatPr'].attributes['defaultColWidth']
      cols = sheet.elements['worksheet/cols']
      # TODO
      if cols && (cols.size > 0 || !width) && cols.elements['col'].attributes['customWidth'] == "1"
          width = sheet.elements['worksheet/cols/col'].attributes['width']
      end

      @default_width = (width.to_f * 7 + 4) * 72 / 96
      @default_height = sheet.elements['worksheet/sheetFormatPr'].attributes['defaultRowHeight'].to_f
      @row_height = []

      cells = []
      sheet.elements.each('/worksheet/sheetData/row') {|row|
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
      @row_pos_y =  @row_height.each_with_object([0]){|h,a|
        a << a.last + (h || @default_height)
      }
      cells
    end

    def col_x(c)
        c * @default_width
    end

    def row_y(r)
        @row_pos_y[r] || r * @default_height
    end

    def load_drawing
      elem = @doc.elements['/worksheet/drawing'] or return []
      rels = @book.read_rels(@sheet)
      path = rels[elem.attributes['r:id']] or return []
      path = (Pathname(@sheet).dirname + path).to_s

      ent = @zip_file.get_entry(path)
      drawing = REXML::Document.new(ent.get_input_stream.read)
      scale = @scale

      drawings = []
      drawing.elements.each('xdr:wsDr/xdr:twoCellAnchor') {|cell_anchor|
        prstGeom = cell_anchor.elements['xdr:sp/xdr:spPr/a:prstGeom']
        next unless prstGeom
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
        geom = prstGeom.attributes['prst']
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
        elsif geom == 'rect'
        end

        drawings << {type: 'two_cell', from_col: from_col, from_row: from_row, height: h, width: w, style: style, text: s}
      }

      # TODO:
      #drawing.elements.each('xdr:wsDr/xdr:oneCellAnchor') {|cell_anchor|
      #  puts "oneCellAnchor:" + cell_anchor.elements['//a:prstGeom'].attributes['prst']
      #   drawings << {type: 'one_cell', col: col, row: row, style: style, text: s}
      #}

      drawings
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
           x = ((cell[:col]-1) * default_width * scale).round(2)
           y = (row_y(cell[:row]-1) * scale).round(2)
           w = ((cell[:w] || 1) * default_width * scale).round(2)
           h = (row_y(cell[:row] + (cell[:h] || 1) - 1) * scale).round(2) - y
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

