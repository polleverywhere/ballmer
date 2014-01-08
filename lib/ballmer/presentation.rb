module Ballmer
  # Represents a presentation that has many slides.
  class Presentation < Document
    autoload :Slides,       'ballmer/presentation/slides'
    autoload :Slide,        'ballmer/presentation/slide'
    autoload :Notes,        'ballmer/presentation/notes'
    autoload :NotesParser,  'ballmer/presentation/notes_parser'
    autoload :Tags,         'ballmer/presentation/tags'
    
    # Return an array of slides.
    def slides
      @slides ||= Slides.new(self)
    end

    # Presentation XML file.
    def presentation
      XMLPart.new(self, '/ppt/presentation.xml')
    end
  end
end
