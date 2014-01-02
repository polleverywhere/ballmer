module Ballmer
  class Presentation
    # Manages concerns around keeping slide and notesSlides files in
    # sync with an array of slides. These basically needs to trasnact
    # the slide\d+ and slideNote\d+ numbers to be in sync with an array.
    # Its a big, ugly ass complicated beast. Send your thank you cards to Bill Gates.
    class Slides
      include Enumerable

      def initialize(doc)
        @doc = doc
      end

      def each(&block)
        # TODO - Do NOT read content-types, but read Rels instead (and move this type casting in there.)
        @doc.content_types[Slide::CONTENT_TYPE].each { |path| block.call slide path }
      end

      # This method is crazy because it has to manipulate a ton of files within the PPTX. Most of
      # what happens in here I figured out by diff-ing PPTX files that had copies of identical slides, but
      # a different number of slides.
      def push(slide)
        n = to_a.size + 1
        # Paths within the zip file of new files we have to write.
        slide_path = Pathname.new("/ppt/slides/slide#{n}.xml")
        slide_rels_path = Pathname.new("/ppt/slides/_rels/slide#{n}.xml.rels")
        slide_notes_path = Pathname.new("/ppt/notesSlides/notesSlide#{n}.xml")
        slide_notes_rels_path = Pathname.new("/ppt/notesSlides/_rels/notesSlide#{n}.xml.rels")
        presentation_rels_path = Pathname.new("/ppt/_rels/presentation.xml.rels")
        presentation_path = Pathname.new("/ppt/presentation.xml")

        # Update ./ppt
        #   !!! CREATE !!!
        #   ./slides
        #     Create new files
        #       ./slide(\d+).xml file
        @doc.copy slide_path, slide.path
        #       ./_rels/slide(\d+).xml.rels
        @doc.copy slide_rels_path, slide.rels.path
        #   ./notesSlides
        #     Create new files
        #       ./notesSlide(\d+).xml file
        @doc.copy slide_notes_path, slide.notes.path
        #       ./_rels/notesSlide(\d+).xml.rels
        @doc.copy slide_notes_rels_path, slide.notes.rels.path
        
        #   !!! UPDATES !!!
        # Update the notes in the new slide to point at the new notes
        @doc.edit_xml slide_rels_path do |xml|
          # TODO - Move this rel logic into the parts so that we don't have to repeat ourselves when calculating this stuff out.
          xml.at_xpath("//xmlns:Relationship[@Type='#{Notes::REL_TYPE}']")['Target'] = slide_notes_path.relative_path_from(slide_path.dirname)
        end

        # Update teh slideNotes reference to point at the new slide
        @doc.edit_xml slide_notes_rels_path do |xml|
          xml.at_xpath("//xmlns:Relationship[@Type='#{Slide::REL_TYPE}']")['Target'] = slide_path.relative_path_from(slide_notes_path.dirname)
        end

        #   ./_rels/presentation.xml.rels
        #     Update Relationship ids
        #     Insert a new one slideRef
        @doc.edit_xml presentation_rels_path do |xml|
          # Calucate the next id
          next_id = xml.xpath('//xmlns:Relationship[@Id]').map{ |n| n['Id'] }.sort.last.succ
          # TODO - Figure out how to make this more MS idiomatic up 9->10 instead of incrementing
          # the character....
          # Insert that into the slide and crakc open the presentation.xml file
          types = xml.at_xpath('/xmlns:Relationships')
          types << Nokogiri::XML::Node.new("Relationship", xml).tap do |n|
            n['Id'] = next_id
            n['Type'] = Slide::REL_TYPE
            n['Target'] = slide_path.relative_path_from(presentation_path.dirname)
          end
          #   ./presentation.xml
          #     Update attr
          #       p:notesMasterId
          #     Insert attr
          #       p:sldId, increment, etc.
          @doc.edit_xml '/ppt/presentation.xml' do |xml|
            slides = xml.at_xpath('/p:presentation/p:sldIdLst')
            next_slide_id = slides.xpath('//p:sldId[@id]').map{ |n| n['id'] }.sort.last.succ
            slides << Nokogiri::XML::Node.new("p:sldId", xml).tap do |n|
              # TODO - Fix the ID that's jacked up.
              n['id'] = next_slide_id
              n['r:id'] = next_id
            end
          end
        end

        # Update ./[Content-Types].xml with new slide link and slideNotes link
        @doc.edit_xml Document::ContentTypes::PATH do |xml|
          types = xml.at_xpath('/xmlns:Types')
          types << Nokogiri::XML::Node.new("Override", xml).tap do |n|
            n['PartName'] = slide_path
            n['ContentType'] = Slide::CONTENT_TYPE
          end
          types << Nokogiri::XML::Node.new("Override", xml).tap do |n|
            n['PartName'] = slide_notes_path
            n['ContentType'] = Notes::CONTENT_TYPE
          end
        end

        # Great, that's all done, so lets return the slide eh?
        slide slide_path
      end

      private
      def slide(path)
        Slide.new(@doc, path) 
      end
      # Reads from the [Content_Types].xml file the paths for the slide
      #   <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
      def parts
        @doc.content_types.parts Slide::CONTENT_TYPE
      end

      # Microsoft decided it would be cool to start at 1 instead of 0 
      # for the part indices, so this deals with that seperatly
      def next_number
        self.to_a.size + 1
      end
    end
  end
end
