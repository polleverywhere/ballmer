module Ballmer
  # Deals with file concerns between higher-level classes like
  # Slides, Notes and file-system level work.
  class Document
    autoload :Archive,      'ballmer/document/archive'
    autoload :Part,         'ballmer/document/part'
    autoload :XMLPart,      'ballmer/document/xml_part'
    autoload :Rels,         'ballmer/document/rels'
    autoload :ContentTypes, 'ballmer/document/content_types'

    # Forward method calls on document to the archive, mostly
    # low-level read/write/copy file operations. The Document
    # class should deal with decorating these read/writes with
    # helper Parts.
    extend Forwardable
    def_delegators :archive, :read, :write, :copy, :delete, :commit

    # Make the archive available in case the developer needs to perform
    # even lower level IO functions.
    attr_reader :archive

    def initialize(archive)
      @archive = archive
    end

    # Open an XML office file from the given path.
    def self.open(path)
      new Archive.open(path)
    end

    # Read zip data from a bufffer. Very useful when you want to load a template 
    # into a server environment, modify, and serve up without writing to disk.
    def self.read(data)
      new Archive.read(data)
    end

    # Open an XML document at the given path.
    def xml(path)
      Nokogiri::XML read path
    end

    # Modify XML within a block and write it back to the zip when done.
    def edit_xml(path, &block)
      write path, xml(path).tap(&block).to_s
    end

    # Content types XML file.
    def content_types
      ContentTypes.new(self)
    end
  end
end
