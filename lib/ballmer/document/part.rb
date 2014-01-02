module Ballmer
  class Document
    # Basic behavior of a part that we lift off of the [Content_Types].xml file.
    class Part
      attr_reader :path, :doc

      def initialize(doc, path)
        @doc, @path = doc, Pathname.new(path)
      end

      def xml
        @xml ||= doc.xml(@path)
      end

      # Grab the rels file for this asset.
      def rels
        Rels.relative_to(self)
      end

      # Commit the part XML to the buffer.
      def commit
        @doc.write path, xml.to_s
      end
    end
  end
end