require 'spec_helper'

describe PowerPoint do
  it "should open" do
    PowerPoint::Presentation.open("Presentation3.pptx").send(:slide_parts).should have(3).items
  end
end