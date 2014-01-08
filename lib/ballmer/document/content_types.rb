module Ballmer
  class Document
    # Deals with everything related to content paths.
    class ContentTypes < XMLPart
      PATH = "[Content_Types].xml"

      def initialize(doc, path = PATH)
        super doc, path
      end

      # Get all of the parts for a given type
      # TODO - Have this return an enumerable of parts so we can fitler by part-type.
      # We can filter this by type from whatever is calling it if we blow open some
      # new types...
      def parts(type)
        xml.xpath("//xmlns:Override[@ContentType='#{type}']").map{ |n| n['PartName'] }
      end
      alias :[] :parts

      # Append a part to ContentTypes
      def append(part)
        # Don't write the part again if it already exists ya dummy
        return nil if exists? part

        edit_xml do |xml|
          xml.at_xpath('/xmlns:Types').tap do |types|
            types << Nokogiri::XML::Node.new("Override", xml).tap do |n|
              n['PartName'] = part.path
              n['ContentType'] = part.class::CONTENT_TYPE
            end
          end
        end
      end
      alias :<< :append

      # Test if the part already exists so we don't write
      # it multiple times.
      def exists?(part)
        !! xml.at_xpath("//xmlns:Override[@ContentType='#{part.class::CONTENT_TYPE}' and @PartName='#{part.path}']")
      end
    end
  end
end
