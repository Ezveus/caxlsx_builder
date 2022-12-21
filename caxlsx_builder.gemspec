# frozen_string_literal: true

require_relative 'lib/caxlsx_builder/version'

Gem::Specification.new do |spec|
  spec.name = 'caxlsx_builder'
  spec.version = CaxlsxBuilder::VERSION
  spec.authors = ['Matthieu CIAPPARA']
  spec.email = ['matthieu.ciappara@outlook.fr']

  spec.summary = 'Write Excel files with ease!'
  spec.description = 'Write Excel files with ease using a DSL'
  spec.homepage = 'https://github.com/Ezveus/caxlsx_builder'
  spec.license = 'MIT'
  spec.required_ruby_version = '>= 3.1.0'

  spec.metadata['homepage_uri'] = spec.homepage
  spec.metadata['source_code_uri'] = spec.homepage
  spec.metadata['changelog_uri'] = 'https://github.com/Ezveus/caxlsx_builder/blob/master/CHANGELOG.md'

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  spec.files = Dir.chdir(__dir__) do
    `git ls-files -z`.split("\x0").reject do |f|
      (f == __FILE__) || f.match(%r{\A(?:(?:bin|test|spec|features)/|\.(?:git|travis|circleci)|appveyor)})
    end
  end
  spec.bindir = 'exe'
  spec.executables = spec.files.grep(%r{\Aexe/}) { |f| File.basename(f) }
  spec.require_paths = ['lib']

  spec.add_dependency 'caxlsx', '~> 3.0'
  spec.metadata['rubygems_mfa_required'] = 'true'
end
