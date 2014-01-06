module Ballmer
  class Presentation
    class NotesParser
      # Make this read the language set by the OS or from a configuration.
      DEFAULT_LANG = 'en-US'

      def initialize(xml)
        @xml = xml
      end

      # MSFT Thought it would be cool to drop a bunch of different bodies
      # and id attributes in here that don't link to anything, so lets go
      # loosey goosey on it and find the stupid "Notes Placeholder" content.
      def node
        @xml.at_xpath('
          //p:nvSpPr[
            p:cNvPr[starts-with(@name, "Notes Placeholder")]
          ]/following-sibling::p:txBody
        ')
      end

      def to_s
        node.xpath('.//a:t').map(&:text).join("\n\n")
      end

      # Parses a text file with newline breaks into "paragraphs" per whatever weird markup
      # the noteSlides is using. For now we're keeping this simple, no italics or other crazy stuff,
      # but this is the class that would be extended, changed, or swapped out in the future.
      def parse(body, lang = DEFAULT_LANG)
        body_pr = Nokogiri::XML::Node.new("p:txBody", @xml)
        # These should be blank... I don't know why, but they're always that way in the files.
        #   <a:bodyPr/>
        #   <a:lstStyle/>
        body_pr << Nokogiri::XML::Node.new("a:bodyPr", @xml)
        body_pr << Nokogiri::XML::Node.new("a:lstStyle", @xml)
        # TODO - Reject blank lines after we chomp 'em
        body_pr << Nokogiri::XML::Node.new("a:p", @xml).tap do |p|
          p << Nokogiri::XML::Node.new("a:r", @xml).tap do |r|
            r << Nokogiri::XML::Node.new("a:rPr", @xml).tap do |rpr|
              # PPT just wants this, k?
              rpr["lang"] = lang
              rpr["dirty"] = "0"
              rpr["smtClean"] = "0"
            end
            # This is where we finally inject content. w00.
            r << Nokogiri::XML::Node.new("a:t", @xml).tap do |t|
              t.content = body
            end
          end
        end
        node.replace body_pr
      end
    end
  end
end
