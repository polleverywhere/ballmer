module Ballmer
  class Document
    # Basic behavior of a part. Whats a part? Could be an image, an XML file, or anything really. Most
    # parts in a document will be an XMLPart, so be sure to take a look at that.
    class Part
      attr_reader :path, :doc

      def initialize(doc, path)
        @doc, @path = doc, Pathname.new(path)
      end

      # Get the relative path for this part from another part. This funciton
      # is mostly used by the Rel class to figure out relationships between parts.
      def relative_path_from(part)
        # I think the rel_part.path bidness is not returning and absolute path. Fix and maybe this will work (and
        # the weird + '..' won't be needed).
        part.path.relative_path_from(path + '..')
      end

      # Commit the part to the buffer.
      def commit(data = self.doc)
        doc.write path, data
      end
    end
  end
end