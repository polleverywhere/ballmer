module Ballmer
  # Represents a presentation that has many slides.
  class Presentation < Document
    # Return an array of slides.
    def slides
      @slides ||= Slides.new(self)
    end
  end

  class Presentation
    autoload :Slides,       'ballmer/presentation/slides'
    autoload :Slide,        'ballmer/presentation/slide'
    autoload :Notes,        'ballmer/presentation/notes'
    autoload :NotesParser,  'ballmer/presentation/notes_parser'
  end
end