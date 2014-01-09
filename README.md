# Ballmer

[![Build Status](https://travis-ci.org/polleverywhere/ballmer.png?branch=master)](https://travis-ci.org/polleverywhere/ballmer) [![Code Climate](https://codeclimate.com/repos/52ce169f6956805d89000475/badges/e90c7084bdc02bda6000/gpa.png)](https://codeclimate.com/repos/52ce169f6956805d89000475/feed)

The Ballmer gem provides the basis for modifying Office documents in Ruby. It provides access to low-level primitives including:

* Unzip/zip Office document formats
* Low level "part" abstraction and "rels" resolution
* Direct access to manipulating/munging XML

PowerPoint is the only format with a higher-level abstraction that allows:

* Copying and inserting slides
* Reading slide notes in the most basic sense
* Writing to slidenotes via a subset of markdown (only paragraphs)

While Word and Excel don't have these abstractions, Ballmer has a "Document" class that can still be used to resolve and manipulate document parts.

## Installation

Add this line to your application's Gemfile:

    gem 'ballmer'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ballmer

## Usage

Its highly recommended to get comfortable with [Nokogiri](http://nokogiri.org) and XPath queries. Here's an example of what Ballmer can do:

```ruby
# Open a pptx file
p = Ballmer::Presentation.open("./fixtures/Presentation3.pptx")
# Copy the first slide into the last position
p.sides.push p.slides.first
# Lets manipulate some XML using XPath queries and Nokogiri.
p.edit_xml 'docProps/app.xml' do |xml|
  xml.at_xpath('/xmlns:Properties/xmlns:Company').content = 'Acme Inc.'
end
# Now save the file.
p.save
```

## Contributing

Microsoft Office documents are a complicating beast. I don't intend no support all functionality, but I do think there's a lot of value in higher-level abstractions for various document formats. If you are working on a project and build these abstractions I'd love to merge those with Ballmer.

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request

## Helpful Information

* [PresentationML](http://msdn.microsoft.com/en-us/library/office/gg278335.aspx) - Directory and XML structure of an Office file.
