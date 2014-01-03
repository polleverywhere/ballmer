module Ballmer
  class Document
    # Manages collections of parts, like slidesN.xml, slidesNotesN.xml, etc.
    class Collection
      include Enumerable

      def initialize(doc)
        @doc = doc
      end

      def each(&block)
      end
    end
  end
end
