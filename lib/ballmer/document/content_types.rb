module Ballmer
  class Document
    # Deals with everything related to content paths.
    class ContentTypes < Part
      PATH = "[Content_Types].xml"

      def initialize(doc, path = PATH)
        super doc, path
      end

      # Get all of the parts for a given type
      # TODO - Have this return an enumerable of parts so we can fitler by part-type.
      def parts(type)
        xml.xpath("//xmlns:Override[@ContentType='#{type}']").map{ |n| n['PartName'] }
      end
      alias :[] :parts
    end
  end
end
