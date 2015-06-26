# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'asposewordsjavaforruby/version'

Gem::Specification.new do |spec|
  spec.name          = 'asposewordsjavaforruby'
  spec.version       = Asposewordsjavaforruby::VERSION
  spec.authors       = ['Aspose Marketplace']
  spec.email         = ['marketplace@aspose.com']
  spec.summary       = %q{A Ruby gem to work with aspose.words libraries}
  spec.description   = %q{AsposeWordsJavaforRuby is a Ruby gem that can help working with aspose.words libraries}
  spec.homepage      = 'https://github.com/asposemarketplace/Aspose_Words_Java_for_Ruby'
  spec.license       = 'MIT'

spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
spec.require_paths = ['lib']

  spec.add_development_dependency 'bundler', '~> 1.7'
  spec.add_development_dependency 'rake', '~> 10.0'
  spec.add_development_dependency 'rspec'

  spec.add_dependency 'rjb', '~> 1.5.3'

end
