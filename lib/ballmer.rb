require "ballmer/version"
require "nokogiri"

module Ballmer
  autoload :Presentation, 'ballmer/presentation'
  autoload :PPTX,         'ballmer/pptx'

  # Deals with everything related to content paths.
  class ContentTypes
    PATH = "[Content_Types].xml"

    attr_reader :path, :pptx

    def initialize(pptx, path = PATH)
      @pptx, @path = pptx, path
    end

    # Get all of the parts for a given type
    # TODO - Have this return an enumerable of parts so we can fitler by part-type.
    def parts(type)
      xml.xpath("//xmlns:Override[@ContentType='#{type}']").map{ |n| n['PartName'] }
    end
    alias :[] :parts

    def xml
      pptx.xml(path)
    end
  end

  # Basic behavior of a part that we lift off of the [Content_Types].xml file.
  class Part
    attr_reader :path, :pptx

    def initialize(pptx, path)
      @pptx, @path = pptx, Pathname.new(path)
    end

    def xml
      @xml ||= pptx.xml(@path)
    end

    # Grab the rels file for this asset.
    def rels
      Rels.relative_to(self)
    end

    # Commit the part XML to the buffer.
    def commit
      @pptx.write path, xml.to_s
    end
  end

  # CRUD and resolve relative documents to a part. These map to .xml.rel documents
  # in the MS Office document format.
  class Rels
    attr_reader :path, :pptx

    def initialize(pptx, path, part_path)
      @pptx, @path, @part_path = pptx, path, part_path
    end

    # Return a list of target paths given a type.
    def targets(type)
      xml.xpath("//xmlns:Relationship[@Type='#{type}']").map{ |n| Pathname.new(n['Target']) }
    end

    def xml
      pptx.xml(@path)
    end

    # Create a Rels class from a given part.
    def self.relative_to(part)
      new part.pptx, rels_path(part.path), part.path
    end

    # Resolve the default rels asset for a given part path.
    def self.rels_path(part_path)
      Pathname.new(part_path).join('../_rels', Pathname.new(part_path).sub_ext('.xml.rels').basename)
    end
  end
end