require 'sketchup.rb'
require 'extensions.rb'

module LSV_To_Excel
  unless file_loaded?(__FILE__)
    path = File.join(File.dirname(__FILE__), 'LSV_To_Excel', 'main.rb')
    ex = SketchupExtension.new('LSV To Excel', path)
    ex.description = 'Export direct de mesures vers Excel avec reglages personnalises.'
    ex.version     = '3.1.0'
    ex.creator     = 'AP Bureau Etude'
    Sketchup.register_extension(ex, true)
    file_loaded(__FILE__)
  end
end