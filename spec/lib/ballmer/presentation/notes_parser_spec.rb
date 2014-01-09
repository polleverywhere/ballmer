require 'spec_helper'

describe Ballmer::Presentation::NotesParser do
  # Fixture contains a complex note with several different formats.
  subject           { read_presentation('notes.pptx').slides.first.notes }

  it "should have notes" do
    subject.body.should == %[Some crazy notes are here. Some words are in bold, others are italic.

Funny thing about notes. When there’s a mispullin’ it gets underlined and tokenized.]
  end

  it "should edit notes" do
    subject.body = "Hi\bpig"
    subject.body.should == "Hi\bpig"
  end
end