module CaxlsxBuilder

  class Cell

    def initialize(style: :default, type: :string)
      @style = style
      @type  = type
    end

    def as_style(item)
      style = @style.respond_to?(:call) ? @style.call(item) : @style
      type  = @type.respond_to?(:call) ? @type.call(item) : @type

      { style:, type: }
    end

  end

end
