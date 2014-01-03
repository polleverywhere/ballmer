module Ballmer
  class Presentation
    class Tags < Document::Part
      # TODO, there are three types of notes. We need to figure out 
      # how to resolve the slide number, notes, and whatever the hell else
      # the first note type is.

      # Key used to look up notes from [Content-Types].xml.
      CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml".freeze

      # Key used to look up notes from .xml.rel documents
      REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'.freeze

      # Read tag
      def [](key)
        if tag = tag(key)
          tag['val']
        end
      end

      # If the tag exists, create it -- otherwise update.
      #   <p:tag name="" val=""/>
      def []=(key, value)
        unless tag key # Don't add the tag if it already exists.
          tag_list << Nokogiri::XML::Node.new("p:tag", @xml) do |tag|
            tag['name'] = key
            tag['val'] = value
          end
        end
      end

      private

      def tag(key)
        tag_list.at_xpath("./p:tag[@name='#{key}']")
      end

      def tag_list
        xml.at_xpath('p:tagLst')
      end

      def self.build(doc, path)
        # Write the template
        doc.write path, File.read(File.join(File.dirname(__FILE__), 'tags.xml'))
        # K, now lets read it and decorate with the part.
        new doc, path
      end
    end
  end
end
