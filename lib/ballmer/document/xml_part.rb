module Ballmer
  class Document
    # Various helpers for editing Part XML data and resolving relative part paths.
    class XMLPart < Part
      # TODO - Figure out how to curry the path into this call and delegate.
      # Also, if the DOM is what speaks the truth, this caching will cause some
      # really stupid/weird bug down the line.
      def xml
        @xml ||= doc.xml(path)
      end

      # TODO - Figure out how to curry the path into this call and delegate
      def edit_xml(&block)
        block.call(xml)
        commit
      end

      # Grab the rels file for this asset.
      def rels
        Rels.from(self)
      end

      # Commit the part XML to the buffer.
      def commit
        super xml.to_s
      end
    end
  end
end
