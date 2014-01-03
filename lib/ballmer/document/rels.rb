module Ballmer
  class Document
    # CRUD and resolve relative documents to a part. These map to .xml.rel documents
    # in the MS Office document format.
    class Rels < Part
      attr_reader :path, :doc

      def initialize(doc, path, part_path)
        super doc, path
        @part_path = part_path
      end

      # Return a list of target paths given a type.
      def targets(type)
        xml.xpath("//xmlns:Relationship[@Type='#{type}']").map{ |n| Pathname.new(n['Target']) }
      end

      # Create a Rels class from a given part.
      def self.relative_to(part)
        new part.doc, rels_path(part.path), part.path
      end

      # Resolve the default rels asset for a given part path.
      def self.rels_path(part_path)
        Pathname.new(part_path).join('../_rels', Pathname.new(part_path).sub_ext('.xml.rels').basename)
      end
    end
  end
end
