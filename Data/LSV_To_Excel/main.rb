require 'sketchup.rb'

module LSV_To_Excel
  module Extension

    # Récupération des réglages
    def self.get_settings
      @precision = Sketchup.read_default("LSV_Excel", "precision", 2)
      @unit      = Sketchup.read_default("LSV_Excel", "unit", "Metrique")
      @separator = Sketchup.read_default("LSV_Excel", "separator", ",")
    end

    # Fenêtre de réglages
    def self.show_settings
      self.get_settings
      prompts = ["Decimales:", "Unite:", "Separateur:"]
      list    = ["", "Metrique|Inch", ",|."]
      defaults = [@precision, @unit, @separator]
      
      input = UI.inputbox(prompts, defaults, list, "Configuration LSV_To_Excel")
      
      if input
        @precision, @unit, @separator = input
        Sketchup.write_default("LSV_Excel", "precision", @precision)
        Sketchup.write_default("LSV_Excel", "unit", @unit)
        Sketchup.write_default("LSV_Excel", "separator", @separator)
      end
    end

    # Logique d'envoi vers Excel
    def self.send_to_excel
      self.get_settings
      selection = Sketchup.active_model.selection
      if selection.empty?
        UI.messagebox("Veuillez selectionner un objet.")
        return
      end

      total_length = 0; total_area = 0; total_volume = 0
      has_vol = false; has_area = false

      selection.each do |ent|
        if ent.is_a?(Sketchup::Edge)
          total_length += ent.length
        elsif ent.is_a?(Sketchup::Face)
          total_area += ent.area; has_area = true
        elsif (ent.is_a?(Sketchup::Group) || ent.is_a?(Sketchup::ComponentInstance))
          if ent.respond_to?(:volume) && ent.volume > 0
            total_volume += ent.volume; has_vol = true
          end
        end
      end

      ratio = 1.0
      if @unit == "Metrique"
        ratio = has_vol ? (0.0254**3) : (has_area ? (0.0254**2) : 0.0254)
      end

      val_num = (has_vol ? total_volume : (has_area ? total_area : total_length)) * ratio
      formatted_val = sprintf("%.#{@precision}f", val_num).gsub('.', @separator)

      begin
        require 'win32ole'
        excel = WIN32OLE.connect("Excel.Application")
        excel.ActiveCell.Value = formatted_val
      rescue
        UI.messagebox("Excel doit etre ouvert avec une cellule selectionnee.")
      end
    end

    # Création de l'interface (Barre d'outils)
    unless file_loaded?(__FILE__)
      tb = UI::Toolbar.new "LSV To Excel"
      path_dir = File.dirname(__FILE__)
      
      # BOUTON EXPORT
      cmd_run = UI::Command.new("Exporter vers Excel") { self.send_to_excel }
      cmd_run.tooltip = "Envoyer la mesure vers Excel"
      icon_export = File.join(path_dir, "icon.png")
      if File.exist?(icon_export)
        cmd_run.small_icon = icon_export
        cmd_run.large_icon = icon_export
      end
      
      # BOUTON RÉGLAGES (ENGRENAGE)
      cmd_set = UI::Command.new("Reglages") { self.show_settings }
      cmd_set.tooltip = "Modifier les preferences"
      icon_reglages = File.join(path_dir, "iconR.png")
      if File.exist?(icon_reglages)
        cmd_set.small_icon = icon_reglages
        cmd_set.large_icon = icon_reglages
      end
      
      tb.add_item cmd_run
      tb.add_item cmd_set
      tb.show
      
      file_loaded(__FILE__)
    end
  end
end