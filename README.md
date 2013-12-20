# PowerPoint

The PowerPoint gem makes it possible to manipulate PowerPoint files
via Ruby at the most basic level.

## Installation

Add this line to your application's Gemfile:

    gem 'power_point'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install power_point

## Usage

```ruby
# Open a pptx file
p = PowerPoint::Presentation.open("./fixtures/Presentation3.pptx")
# Copy the first slide into the last position
p.sides.push p.slides.first
# Now save the file.
p.save
```

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
