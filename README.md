# Ballmer

The Ballmer gem provides the basis for modifying Office documents in Ruby. It provides access to low-level primitives including:

* Unzip/zip Office document formats.
* Low level "part" abstraction and rels resolution.
* Direct access to manipulating/munging XML.

PowerPoint is the only format with a higher-level, but basic abstraction that allows:

* Copying and inserting slides.
* Reading slide notes in the most basic sense. 
* Writing to slidenotes via a subset of markdown (only paragraphs).

## Installation

Add this line to your application's Gemfile:

    gem 'ballmer'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ballmer

## Usage

```ruby
# Open a pptx file
p = Ballmer::PowerPoint.open("./fixtures/Presentation3.pptx")
# Copy the first slide into the last position
p.sides.push p.slides.first
# Now save the file.
p.save
```

## Contributing

Microsoft Office is a complicating beast.

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request

## Helpful Information

* [PresentationML](http://msdn.microsoft.com/en-us/library/office/gg278335.aspx)
