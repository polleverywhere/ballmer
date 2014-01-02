module Ballmer
  class Presentation
    # Load a slide up in thar. 
    class Slide < Document::Part
      # Key used to look up slides from [Content-Types].xml.
      CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".freeze

      # Key used to look up slides from .xml.rel documents
      REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide".freeze

      def notes
        # TODO - Move a type caster into rels based on content type like 
        # rels[Notes::REL_TYPE].first
        notes_path = rels.targets(Notes::REL_TYPE).first.expand_path(@path.dirname)
        Notes.new(@doc, notes_path)
      end
    end
  end
end
