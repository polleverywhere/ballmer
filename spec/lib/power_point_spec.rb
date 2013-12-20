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

    it "should have notes" do
      slide.notes.body.should =~ /Poll A/
    end
  end

  context "slides" do
    let(:original)  { subject.slides.to_a.first }
    let(:copy)      { subject.slides.push original }

    it "should copy slide" do
      original.notes.body.should == copy.notes.body
    end
  end
end