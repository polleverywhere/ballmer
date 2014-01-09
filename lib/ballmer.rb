require "ballmer/version"
require "nokogiri"

module Ballmer
  autoload :Presentation, 'ballmer/presentation'
  autoload :Document,     'ballmer/document'

  # Get an absolute path relative to ballmer.
  def self.path(*args)
    File.expand_path(File.join('..', args), __FILE__)
  end
end