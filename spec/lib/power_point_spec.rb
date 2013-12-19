require 'spec_helper'

describe PowerPoint do
  subject { PowerPoint::Presentation.open("./fixtures/Presentation3.pptx") }

  it "should have slides" do
    subject.slides.should have(3).items
  end

  context "slide" do
    let(:slide) { subject.slides.first }

    it "should have notes" do
      slide.notes.should =~ /Poll A/
    end
  end
end