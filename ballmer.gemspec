# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'ballmer/version'

Gem::Specification.new do |spec|
  spec.name          = "ballmer"
  spec.version       = Ballmer::VERSION
  spec.authors       = ["Brad Gessler"]
  spec.email         = ["brad@polleverywhere.com"]
  spec.description   = %q{Open and manipulate Ballmer files.}
  spec.summary       = %q{Manipulate Ballmer slides in Ruby.}
  spec.homepage      = ""
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "guard-rspec"
  spec.add_development_dependency "terminal-notifier-guard"
  spec.add_development_dependency "pry"
  
  spec.add_dependency "zipruby"
  spec.add_dependency "nokogiri"
end
