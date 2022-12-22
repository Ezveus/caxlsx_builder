# frozen_string_literal: true

RSpec.describe CaxlsxBuilder do
  it "has a version number" do
    expect(CaxlsxBuilder::VERSION).not_to be nil
  end

  describe CaxlsxBuilder::Builder do
    it "can generate a simple file with one string column, a header row and two data rows" do
      package = CaxlsxBuilder::Builder.new({ 'Test' => [['a', 1.99, 'Yay!'], ['b', 2.015, 'Oh?']] }) { |sheet|
        # Add a column with default style and no footer
        sheet.column('Interjections') { |item| item.last }
      }.call

      generated_file = RubyXL::Parser.parse_buffer(package.to_stream)

      # Check the number of sheets
      expect(generated_file.worksheets.size).to eq(1)

      # Check the name of the sheet
      sheet = generated_file.worksheets.first
      expect(sheet.sheet_name).to eq('Test')

      # First row is the headers
      expect(sheet[0][0].value).to eq('Interjections')
      expect(sheet[0][1]).to be_nil

      # Second row is the interjection 'Yay!'
      expect(sheet[1][0].value).to eq('Yay!')
      expect(sheet[1][1]).to be_nil

      # Third row is the interjection 'Oh?'
      expect(sheet[2][0].value).to eq('Oh?')
      expect(sheet[2][1]).to be_nil

      # Fourth row is empty
      expect(sheet[3][0].value).to be_nil
      expect(sheet[3][1]).to be_nil
    end

    it "can generate a file with three various columns, a header row, two data rows and a footer" do
      package = CaxlsxBuilder::Builder.new({ 'This is my sheet !' => [['a', 1.99, 'Yay!'], ['b', 2.015, 'Oh?']] }) { |sheet|
        # Add a first column with a simple footer
        sheet.column('Letters', footer: 'Some letters') { |item| item.first }

        # Add a second column with some dynamic style, a float type and a dynamic footer
        sheet.column('Numbers',
                     style:  ->(item) { item[1] > 2 ? :header : :default },
                     footer: ->(cells) { cells.sum }, type: :float) { |item| item[1] }

        # Add a last column with default style and no footer
        sheet.column('Interjections') { |item| item.last }
      }.call

      generated_file = RubyXL::Parser.parse_buffer(package.to_stream)

      # Check the number of sheets
      expect(generated_file.worksheets.size).to eq(1)

      # Check the name of the sheet
      sheet = generated_file.worksheets.first
      expect(sheet.sheet_name).to eq('This is my sheet !')

      # First row is the headers
      expect(sheet[0][0..3].map(&:value)).to eq(%w[Letters Numbers Interjections])
      # Header so we expect the header style on this cell
      expect(sheet[0][0..3].map { |cell| cell.font_size }.uniq).to eq([16.0])
      expect(sheet[0][0..3].map { |cell| cell.get_cell_font.b&.val }.uniq).to eq([true])

      # Second row is the interjection 'Yay!'
      expect(sheet[1][0..3].map(&:value)).to eq(['a', 1.99, 'Yay!'])
      # 1.99 < 2 so we expect the normal style on this cell
      expect(sheet[1][0..3].map { |cell| cell.font_size }.uniq).to eq([12.0])
      expect(sheet[1][0..3].map { |cell| cell.get_cell_font.b }.uniq).to eq([nil])

      # Third row is the interjection 'Oh?'
      expect(sheet[2][0..3].map(&:value)).to eq(['b', 2.015, 'Oh?'])
      # 2.015 > 2 so we expect the header style on this cell
      expect(sheet[2][0..3].map { |cell| cell.font_size }).to eq([12.0, 16.0, 12.0])
      expect(sheet[2][0..3].map { |cell| cell.get_cell_font.b&.val }).to eq([nil, true, nil])

      # Fourth row is the footer
      expect(sheet[3][0..3].map(&:value)).to eq(['Some letters', 1.99 + 2.015, nil])
      # Footer so we expect the footer style on this cell
      expect(sheet[3][0..3].map { |cell| cell.font_size }.uniq).to eq([16.0])
      expect(sheet[3][0..3].map { |cell| cell.get_cell_font.b }.uniq).to eq([nil])

      # Fifth row is empty
      expect(sheet[4]).to be_nil
    end

    it "can generate a file with custom styles" do
      package = CaxlsxBuilder::Builder.new({ 'This is my sheet !' => [['a', 1.99, 'Yay!'], ['b', 2.015, 'Oh?']] }) { |sheet|
        # Add a custom style
        sheet.add_style(:blue_on_red, { bg_color: 'FF0000', fg_color: '0000FF' })
        sheet.add_style(:white_on_brown, { bg_color: '936E00', fg_color: 'FFFFFF' })

        # Add a first column with a simple footer and a specific static style
        sheet.column('Letters', footer: 'Some letters', style: :blue_on_red) { |item| item.first }

        # Add a second column with some dynamic style, a float type and a dynamic footer
        sheet.column('Numbers',
                     style:  ->(item) { item[1] > 2 ? :white_on_brown : :default },
                     footer: ->(cells) { cells.sum }, type: :float) { |item| item[1] }

        # Add a last column with default style and no footer
        sheet.column('Interjections') { |item| item.last }
      }.call

      generated_file = RubyXL::Parser.parse_buffer(package.to_stream)
      sheet          = generated_file.worksheets.first

      # First cell should be blue on red
      # 1.99 < 2 so we expect the normal style on the second cell
      # Third is normal
      expect(sheet[1][0..3].map { |cell| cell.font_size }).to eq([11.0, 12.0, 12.0])
      expect(sheet[1][0..3].map { |cell| cell.font_color }).to eq(%w[FF0000FF 000000 000000])
      expect(sheet[1][0..3].map { |cell| cell.fill_color }).to eq(%w[FFFF0000 ffffff ffffff])

      # First cell should be blue on red
      # 2.015 > 2 so we expect white on brown on the second cell
      # Third is normal
      expect(sheet[2][0..3].map { |cell| cell.font_size }).to eq([11.0, 11.0, 12.0])
      expect(sheet[2][0..3].map { |cell| cell.font_color }).to eq(%w[FF0000FF FFFFFFFF 000000])
      expect(sheet[2][0..3].map { |cell| cell.fill_color }).to eq(%w[FFFF0000 FF936E00 ffffff])
    end

    it "can generate a file with many sheets" do
      package = CaxlsxBuilder::Builder.new({ 'A' => [['a']], 'B' => [['b']] }) { |sheet|
        # Add a custom style
        if sheet.name == 'A'
          sheet.add_style(:specific, { bg_color: 'FF0000', fg_color: '0000FF' })
        else
          sheet.add_style(:specific, { bg_color: '936E00', fg_color: 'FFFFFF' })
        end

        # Add a first column with a simple footer and a specific static style
        sheet.column('Letters', style: :specific) { |item| item.first }
      }.call

      generated_file = RubyXL::Parser.parse_buffer(package.to_stream)

      # Check the number of sheets
      expect(generated_file.worksheets.size).to eq(2)

      # Check the name of the sheet
      sheet_a = generated_file.worksheets.first
      expect(sheet_a.sheet_name).to eq('A')
      sheet_b = generated_file.worksheets[1]
      expect(sheet_b.sheet_name).to eq('B')

      # First cell on sheet_a should be blue on red
      expect(sheet_a[1][0..3].map { |cell| cell.font_size }).to eq([11.0])
      expect(sheet_a[1][0..3].map { |cell| cell.font_color }).to eq(%w[FF0000FF])
      expect(sheet_a[1][0..3].map { |cell| cell.fill_color }).to eq(%w[FFFF0000])

      # First cell on sheet_b should be white on brown on the second cell
      expect(sheet_b[1][0..3].map { |cell| cell.font_size }).to eq([11.0])
      expect(sheet_b[1][0..3].map { |cell| cell.font_color }).to eq(%w[FFFFFFFF])
      expect(sheet_b[1][0..3].map { |cell| cell.fill_color }).to eq(%w[FF936E00])
    end
  end
end
