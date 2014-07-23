require 'spec_helper'

describe Ballmer::Presentation do
  subject { read_presentation('presentation3.pptx') }

  context "presentation" do
    it "should have slides" do
      subject.slides.should have(3).items
    end
  end

  it "should have slides#length" do
    subject.slides.length.should == 3
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
            rels = slide.rels.targets(Ballmer::Presentation::Notes::REL_TYPE)
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
            rels = notes.rels.targets(Ballmer::Presentation::Slide::REL_TYPE)
            rels.should have(1).item
            rels.first.to_s.should == '../slides/slide4.xml'
          end
        end
      end

      context "[Content-Type].xml" do
        it "should add slide Override" do
          subject.content_types[Ballmer::Presentation::Slide::CONTENT_TYPE].should have(4).items
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

    context "#delete" do
      before(:each) do
        subject.slides.delete subject.slides.first
      end

      it "should remove slide" do
        subject.should have(3).slides
      end

      context "[Content-Type].xml" do
        it "should remove slide"
      end

      context "presentation.xml" do
        it "should remove slide"
        context "_rels" do
          it "should remove slide"
        end
      end

      context "slide.xml" do
        it "should delete"
        context "rels" do
          it "should delete"
        end
      end

      context "noteSlide.xml" do
        it "should delete"
        context "rels" do
          it "should delete"
        end
      end
    end
  end
end
