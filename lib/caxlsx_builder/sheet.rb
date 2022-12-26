module CaxlsxBuilder

  class Sheet

    attr_reader :name, :headers, :cell_styles, :footers, :styles

    def initialize(name)
      @name          = name
      @headers       = []
      @cell_styles   = []
      @footers       = []
      @cell_builders = []
      @row_cells     = []
      @styles        = {}
    end

    def column(header, style: :default, type: :string, footer: nil, &cell_builder)
      @headers << header
      @cell_styles << Cell.new(style:, type:)
      @footers << footer
      @cell_builders << cell_builder
    end

    def add_style(name, style)
      @styles[name] = style
    end

    def make_row(item, rescue_errors: false)
      # Build the new row
      row = @cell_builders.map do |cell|
        cell.call(item)
      rescue StandardError => e
        return nil if rescue_errors

        raise e
      end

      # Add the new row to the cache then return it
      (@row_cells << row).last
    end

    def cells(index)
      @row_cells.flat_map { |row| row[index] }
    end

  end

end
