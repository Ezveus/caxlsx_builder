# CaxlsxBuilder

Write Excel files with ease using a DSL!

## Installation

Install the gem and add to the application's Gemfile by executing:

    $ bundle add caxlsx_builder

If bundler is not being used to manage dependencies, install the gem by executing:

    $ gem install caxlsx_builder

## Usage

```ruby
# Initialize the builder
builder = CaxlsxBuilder::Builder.new({ 'Test' => [['a', 1.99, 'Yay!'], ['b', 2.015, 'Oh?']] }) do |sheet|
  # Add a custom style
  sheet.add_style(:white_on_brown, { bg_color: '936E00', fg_color: 'FFFFFF' })

  # Add a first column with a simple footer
  sheet.column('Letters', footer: 'Some letters') { |item| item.first }

  # Add a second column with some dynamic style, a float type and a dynamic footer
  sheet.column('Numbers',
               style:  ->(item) { item[1] > 2 ? :white_on_brown : :default },
               footer: ->(cells) { cells.sum }, type: :float) { |item| item[1] }
  
  # Add a last column with default style and no footer
  sheet.column('Interjections') { |item| item.last }
end

# Generate the AXLSX package
package_axlsx = builder.call

# Use Axlsx package's methods to store/write the file

# Write on disk
package_axlsx.serialize('test.xlsx')

# Store with ActiveStorage (not a dependency of this gem): https://guides.rubyonrails.org/active_storage_overview.html
class Export < ApplicationRecord
  has_one_attached :file
end

file   = package_axlsx.to_stream.string
export = Export.new
export.file.attach(io:           StringIO.new(file), filename: 'test.xlsx',
                   content_type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
```

Result:

![Excel export with 3 columns a header two rows and a footer](/test.xlsx.png)

CaxlsxBuilder use the [Caxlsx](https://github.com/caxlsx/caxlsx) gem. The builder directly returns the Caxlsx package to let you do what you want with it.

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/Ezveus/caxlsx_builder. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/Ezveus/caxlsx_builder/blob/master/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the CaxlsxBuilder project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/Ezveus/caxlsx_builder/blob/master/CODE_OF_CONDUCT.md).
