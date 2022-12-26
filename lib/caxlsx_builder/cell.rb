module CaxlsxBuilder

  class Cell

    def initialize(style: :default, type: :string)
      @style = style
      @type  = type
    end

    def as_style(item, rescue_errors: false)
      style = @style.respond_to?(:call) ? call_style_proc(item, rescue_errors:) : @style
      type  = @type.respond_to?(:call) ? call_type_proc(item, rescue_errors:) : @type

      { style:, type: }
    end

    private

    def call_style_proc(item, rescue_errors: false)
      @style.call(item)
    rescue StandardError => e
      if rescue_errors
        :default
      else
        raise e
      end
    end

    def call_type_proc(item, rescue_errors: false)
      @type.call(item)
    rescue StandardError => e
      if rescue_errors
        :string
      else
        raise e
      end
    end

  end

end
