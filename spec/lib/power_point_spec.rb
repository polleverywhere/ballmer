require 'spec_helper'

describe PowerPoint do
  subject { PowerPoint::Presentation.open("./fixtures/Presentation3.pptx") }

  context "presentation" do
    it "should have slides" do
      subject.slides.should have(3).items
    end
  end

  context "slide" do
    let(:slide) { subject.slides.first }
    let(:note)  { "The pig jumped over the fence.\n\nThe pig then squealed." }

    context "notes" do
      it "should have notes" do
        slide.notes.body.should =~ /Poll A/
      end

      it "should edit notes" do
        slide.notes.body = note
        slide.notes.body.should == note
        # slide.notes.xml.at_xpath('//p:txBody/a:p/a:r/a:t').content.should == "The pig jumped over the fence and got hit by a truck."
      end
    end
  end

  context "slides" do
    let(:original)  { subject.slides.to_a.first }
    let(:copy)      { subject.slides.push original }

    # Explicitly call copy before each test in case
    # we don't call "copy" from a spec below since
    # we don't make assertions on "copy" per test.
    before(:each){ copy }

    context "#push" do
      it "should copy slide" do
        original.notes.body.should == copy.notes.body
      end

      context "slide.xml" do
        let(:slide) { copy }

        it "should append" do
          slide.path.to_s.should == '/ppt/slides/slide4.xml'
        end

        context "rels" do
          it "should append ./_rels/slide.xml" do
            slide.rels.path.to_s.should == '/ppt/slides/_rels/slide4.xml.rels'
          end

          it "should reference notesSlides.xml" do
            rels = slide.rels.targets(PowerPoint::Notes::REL_TYPE)
            rels.should have(1).item
            rels.first.to_s.should == '../notesSlides/notesSlide4.xml'
          end
        end
      end

      context "notesSlides.xml" do
        let(:notes) { copy.notes }

        it "should append" do
          notes.path.to_s.should == '/ppt/notesSlides/notesSlide4.xml'
        end

        context "rels" do
          it "should append ./_rels/notesSlides.xml" do
            notes.rels.path.to_s.should == '/ppt/notesSlides/_rels/notesSlide4.xml.rels'
          end

          it "should reference slides.xml" do
            rels = notes.rels.targets(PowerPoint::Slide::REL_TYPE)
            rels.should have(1).item
            rels.first.to_s.should == '../slides/slide4.xml'
          end
        end
      end

      context "[Content-Type].xml" do
        it "should add slide Override" do
          subject.pptx.content_types[PowerPoint::Slide::CONTENT_TYPE].should have(4).items
        end
      end

      context "presentation.xml" do
        it "should add slide"
        it "should add slideId"

        context "_rels" do
          it "should add slide"
          it "should add "
        end
      end
    end
  end

  context "parts" do
    it "should traverse parts" do
      # But why????
      # subject.slides.first.rels[Notes::REL_TYPE].first.
    end
  end
end