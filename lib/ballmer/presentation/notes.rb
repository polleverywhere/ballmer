module Ballmer
  class Presentation
    class Notes < Document::XMLPart
      # TODO, there are three types of notes. We need to figure out 
      # how to resolve the slide number, notes, and whatever the hell else
      # the first note type is.

      # Key used to look up notes from [Content-Types].xml.
      CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml".freeze

      # Key used to look up notes from .xml.rel documents
      REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'.freeze

      # TODO - Generate/update the notes with a mark-down-ish heuristic, 
      # being that two newlines translate into the weird note formats of PPT slides.
      def body=(body)
        nodes_parser.parse(body)
        commit
      end

      def body
        nodes_parser.to_s
      end
      alias :to_s :body

      private
      def nodes_parser
        NotesParser.new(xml)
      end
    end
  end
end
