module CaxlsxBuilder

  class Builder

    CELL_TYPES = %i[date time float integer richtext string boolean iso_8601 text].freeze

    def initialize(sheets, &builder)
      @sheets  = sheets
      @builder = builder
      @styles  = {}
    end

    # @return [Axlsx::Package]
    def call
      # Creating, initializing and returning the +Axlsx+ package
      Axlsx::Package.new do |package|
        package.use_shared_strings = true

        workbook = package.workbook

        workbook.styles do |styles|
          define_default_styles(styles)

          build_sheets(workbook, styles)
        end
      end
    end

    private

    # Define some default styles/formats for the spreadsheet:
    #   default, header, currency and date
    # @param styles [Axlsx::Styles]
    def define_default_styles(styles)
      # Defines the styles of the spreadsheet
      #   default style
      define_style(:default, styles.add_style(sz: 12))
      #   for the cells of the header
      define_style(:header, styles.add_style(b: true, sz: 16))
      #   for the amounts
      define_style(:currency, styles.add_style(format_code: '#.00 &quot;€&quot;'))
      #   for the dates
      define_style(:date, styles.add_style(format_code: 'dd/mm/yyyy'))
      #   for the cells of the footer
      define_style(:footer, styles.add_style(sz: 16))
      #   for wrapped cells
      define_style(:wrapped, styles.add_style(alignment: { wrap_text: true }))
    end

    # Register a custom style under the given name
    # @param name [String,Symbol] A name for the style
    # @param style [Integer] An Excel style identifier
    #   e.g. the result of `workbook.styles { |styles| styles.add_style() }`
    # @return [Symbol] The given name as a Symbol
    def define_style(name, style)
      name          = name.to_sym
      @styles[name] = style

      name
    end

    # @param workbook [Axlsx::Workbook]
    # @param styles [Axlsx::Styles]
    def build_sheets(workbook, styles)
      @sheets.each do |name, collection|
        sheet = Sheet.new(name)
        @builder.call(sheet)

        add_custom_styles(sheet, styles)

        # Add a worksheet
        workbook.add_worksheet(name:) do |worksheet|
          add_header_row(sheet, worksheet)

          add_data_rows(collection, sheet, worksheet)

          add_footer_row(sheet, worksheet)
        end
      end
    end

    # @param sheet [CaxlsxBuilder::Sheet]
    # @param styles [Axlsx::Styles]
    def add_custom_styles(sheet, styles)
      sheet.styles.each do |style_name, style|
        define_style(style_name, styles.add_style(style))
      end
    end

    # @param sheet [CaxlsxBuilder::Sheet]
    # @param worksheet [Axlsx::Worksheet]
    def add_header_row(sheet, worksheet)
      # First row is the head with the header style applied
      worksheet.add_row(sheet.headers, style: @styles[:header])
    end

    # @param collection [#each]
    # @param sheet [CaxlsxBuilder::Sheet]
    # @param worksheet [Axlsx::Worksheet]
    def add_data_rows(collection, sheet, worksheet)
      collection.each do |item|
        row = sheet.make_row(item)

        # Nil row or empty row means no row
        compacted_row = row.compact
        next if compacted_row.nil? || compacted_row.empty?

        # Cells options returns an array which is logic: one option per sheet cell
        # But `sheet.add_row` does not handle arrays correctly, so I transform it to a valid
        # Hash using `cast_to_hash`.
        cells   = sheet.cell_styles.map { |cell| cell_options(**cell.as_style(item)) }
        options = cast_to_hash(cells)
        worksheet.add_row(row, options)
      end
    end

    def add_footer_row(sheet, worksheet)
      return if sheet.footers.empty?

      footer = sheet.footers.map.with_index do |cell, index|
        if cell.respond_to?(:call)
          cell.call(sheet.cells(index))
        else
          cell
        end
      rescue StandardError
        nil
      end

      # Last row is the footer with the footer style applied
      worksheet.add_row(footer, style: @styles[:footer])
    end

    # @return [Hash] The options to apply to a cell
    def cell_options(style: nil, type: nil)
      style = get_style(style)
      type  = get_type(type)

      { style:, type: }
    end

    # Take an array of cell options and map it to a row option
    def cast_to_hash(array)
      return array if array.is_a?(Hash)

      { style: array.map { |i| i[:style] }, types: array.map { |i| i[:type] } }
    end

    # @param style [String,Symbol,nil] The name of the style
    # @return [Integer] The Excel style for the given name or the default cell style
    def get_style(style)
      return @styles[:default] if style.nil? || style == ''

      @styles[style.to_sym] || @styles[:default]
    end

    # @param type [Symbol,nil]
    # @return [Symbol]
    def get_type(type)
      return :string unless CELL_TYPES.include?(type)

      type
    end

  end

end
