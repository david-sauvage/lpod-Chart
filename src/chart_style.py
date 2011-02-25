from lpod.element import register_element_class, odf_create_element, odf_element
from lpod.style import odf_style, odf_create_style


def odf_create_chart_style(name):
    elt = odf_create_element('<style:style style:family="chart"/>')
    elt.set_attribute("style:name", name)
    return elt

class odf_chart_style(odf_style):
    """
    This class implements an xml element that represents a style
    used by charts
    """

###
#Private API
###
    def _set_attribute_in_properties(self, properties, attribute_name, value):
        """
        properties - str
        attribute_name - str
        value - str
        """
        p = self.get_element(properties)
        if p is None:
            p = odf_create_element(properties)
            self.append(p)
        p.set_attribute(attribute_name, value)

    def _get_attribute_in_properties(self, properties, attribute_name):
        """
        properties - str
        attribute_name - str
        """
        p = self.get_element(properties)
        if p is not None:
            return p.get_attribute(attribute_name)
        else:
            return None


#Autoposition       
    def set_auto_position(self, value):
        """value - 'true' 'false'"""
	self._set_attribute_in_properties("style:chart-properties",
                                                  "chart:auto-position", value)

    def get_auto_position(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                         "chart:auto-position")
 

#Autosize
    def set_auto_size(self, value):
        """value - 'true' 'false'"""
	self._set_attribute_in_properties("style:chart-properties",
                                                      "chart:auto-size", value)

    def get_auto_size(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                             "chart:auto-size")

#Font size
    def set_font_size(self, value):
        """value - str like '10pt'"""
	self._set_attribute_in_properties("style:text-properties",
                                                         "fo:font-size", value)

	self._set_attribute_in_properties("style:text-properties",
                                                "style:font-size-asian", value)

	self._set_attribute_in_properties("style:text-properties",
                                              "style:font-size-complex", value)

    def get_font_size(self):
	return self._get_attribute_in_properties("style:text-properties",
                                                                "fo:font-size")

#stroke
    def set_stroke(self, stroke, color):
        """
        stroke - 'none' 'dash' 'solid'
        color - str color in hex like "#b3b3b3"
        """
	self._set_attribute_in_properties("style:graphic-properties",
                                                         "draw:stroke", stroke)

	self._set_attribute_in_properties("style:graphic-properties",
                                                     "svg:stroke-color", color)


    def get_stroke(self):
	return self._get_attribute_in_properties("style:graphic-properties",
                                                                 "draw:stroke")
     
    def get_stroke_color(self):
	return self._get_attribute_in_properties("style:graphic-properties",
                                                            "svg:stroke-color")

#fill
    def set_fill(self, fill, color):
        """
        fill - 'none' 'solid' 'bitmap' 'gradient' 'hatch'
        color - str color in hex like "#b3b3b3"
        """
	self._set_attribute_in_properties("style:graphic-properties",
                                                             "draw:fill", fill)

	self._set_attribute_in_properties("style:graphic-properties",
                                                      "draw:fill-color", color)


    def get_fill(self):
	return self._get_attribute_in_properties("style:graphic-properties",
                                                                   "draw:fill")
     
    def get_fill_color(self):
	return self._get_attribute_in_properties("style:graphic-properties",
                                                             "draw:fill-color")

#include hidden cells
    def set_include_hidden_cells(self, value):
        """value - 'true' 'false'"""
	self._set_attribute_in_properties("style:chart-properties",
                                           "chart:include-hidden-cells", value)

    def get_include_hidden_cells(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                  "chart:include-hidden-cells")

#data style
    def set_data_style(self, style):
        self.set_attribute("style:data-style-name", style)

    def get_data_style(self):
        return self.get_attribute("style:data-style-name")

#display label
    def set_display_label(self, value):
        """value - 'true' 'false'"""
	self._set_attribute_in_properties("style:chart-properties",
                                                 "chart:display-label", value)

    def get_display_label(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                        "chart:display-label")

#logarithmic
    def set_logarithmic(self, value):
        """value - 'true' 'false'"""
	self._set_attribute_in_properties("style:chart-properties",
                                                   "chart:logarithmic", value)

    def get_logarithmic(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                          "chart:logarithmic")

#reverse direction
    def set_reverse_direction(self, value):
        """value - 'true' 'false'"""
	self._set_attribute_in_properties("style:chart-properties",
                                             "chart:reverse-direction", value)

    def get_reverse_direction(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                    "chart:reverse-direction")

#line-break
    def set_line_break(self, value):
        """value - 'true' 'false'"""
	self._set_attribute_in_properties("style:chart-properties",
                                                     "text:line-break", value)

    def get_line_break(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                            "text:line-break")

#axis-position
    def set_axis_position(self, value="0"):
        """value - str like "0" """
	self._set_attribute_in_properties("style:chart-properties",
                                                 "chart:axis-position", value)

    def get_axis_position(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                        "chart:axis-position")

#rotation-angle
    def set_rotation_angle(self, value="0"):
        """value - str like "0" """
	self._set_attribute_in_properties("style:chart-properties",
                                                "style:rotation-angle", value)

    def get_rotation_angle(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                       "style:rotation-angle")

#stroke-width
    def set_stroke_width(self, value):
        """value - str like "0.1cm" """
	self._set_attribute_in_properties("style:graphic-properties",
                                                    "svg:stroke-width", value)

    def get_stroke_width(self):
	return self._get_attribute_in_properties("style:graphic-properties",
                                                           "svg:stroke-width")

#symbol-type
    def set_symbol_type(self, value):
        """value - 'none' 'automatic' 'name-symbol' 'image' """
	self._set_attribute_in_properties("style:chart-properties",
                                                   "chart:symbol-type", value)

    def get_symbol_type(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                          "chart:symbol-type")

#series-source
    def set_series_source(self, value):
        """value - 'rows' 'columns' """
	self._set_attribute_in_properties("style:chart-properties",
                                                 "chart:series-source", value)

    def get_series_source(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                        "chart:series-source")

#symbol-name
    def set_symbol_name(self, value):
        """
        value - 'square' 'diamond' 'arrow-down' 'arrow-up' 'arrow-right' 
                'arrow-left' 'bow-tie' 'hourglass' 'circle' 'star' 'x' 'plus'
                'asterisk' 'horizontal-bar' 'vertical-bar'
        """
	self._set_attribute_in_properties("style:chart-properties",
                                                   "chart:symbol-name", value)

    def get_symbol_name(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                          "chart:symbol-name")

#symbol-width
    def set_symbol_width(self, value):
        """value - str like "0.25cm" """
	self._set_attribute_in_properties("style:chart-properties",
                                                  "chart:symbol-width", value)

    def get_symbol_width(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                         "chart:symbol-width")

#symbol-height
    def set_symbol_height(self, value):
        """value - str like "0.25cm" """
	self._set_attribute_in_properties("style:chart-properties",
                                                 "chart:symbol-height", value)

    def get_symbol_height(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                        "chart:symbol-height")

#treat-empty-cells
    def set_treat_empty_cells(self, value):
        """value - 'leave-gap' """

	self._set_attribute_in_properties("style:chart-properties",
                                             "chart:treat-empty-cells", value)

    def get_treat_empty_cells(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                    "chart:treat-empty-cells")

#right-angled-axes
    def set_right_angled_axes(self, value):
        """value - 'true' 'false'  """
	self._set_attribute_in_properties("style:chart-properties",
                                             "chart:right-angled-axes", value)

    def get_right_angled_axes(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                    "chart:right-angled-axes")

#stacked
    def set_stacked(self, value):
        """value - 'true' 'false' """
	self._set_attribute_in_properties("style:chart-properties",
                                                       "chart:stacked", value)

    def get_stacked(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                              "chart:stacked")

#vertical
    def set_vertical(self, value):
        """value - 'true' 'false' """
	self._set_attribute_in_properties("style:chart-properties",
                                                      "chart:vertical", value)

    def get_vertical(self):
	return self._get_attribute_in_properties("style:chart-properties",
                                                             "chart:vertical")

###
#Quick chart style
###

def odf_create_chart_title_style(name, fontsize):
    """
    Create a simple style to fix a font size

    name - str (name used in chart to use this style)
    fontsize - str (like '12pt')
    """
    s = odf_create_chart_style(name)
    s.set_auto_position("true")
    s.set_font_size(fontsize)
    return s

def odf_create_chart_legend_style(name):
    """
    Create a basic style used for a legend

    name - str (name used in chart to use this style)
    """
    s = odf_create_chart_style(name)
    s.set_auto_position("true")
    s.set_stroke("none", "#b3b3b3")
    s.set_fill("none", "#e6e6e6")
    s.set_font_size("8pt")
    return s

def odf_create_chart_plot_area_style(name):
    """
    Create a basic style for a plot area

    name - str (name used in chart to use this style)
    """
    s = odf_create_chart_style(name)
    s.set_include_hidden_cells("false")
    s.set_auto_size("true")
    s.set_auto_position("true")
    s.set_treat_empty_cells("leave-gap")
    s.set_right_angled_axes("true")
    return s
    
def odf_create_chart_axis_style(name):
    """
    Create a basic style for an axis

    name - str (name used in chart to use this style)
    """
    s = odf_create_chart_style(name)
    s.set_display_label("true")
    s.set_logarithmic("false")
    s.set_reverse_direction("false")
    s.set_line_break("true")
    s.set_axis_position("0")
    s.set_stroke("none", "#b3b3b3")
    s.set_font_size("8pt")
    return s

def odf_create_chart_axis_title_style(name, angle):
    """
    Create a basic style for an axis title

    name - str (name used in chart to use this style)
    """
    
    s = odf_create_chart_style(name)
    s.set_auto_position("true")
    s.set_rotation_angle(angle)
    s.set_font_size("9pt")
    return s

def odf_create_chart_wall_style(name):
    """
    Create a basic style for a wall

    name - str (name used in chart to use this style)
    """
    s = odf_create_chart_style(name)
    s.set_stroke("solid", "#b3b3b3")
    s.set_fill("none", "#e6e6e6")
    return s


#register
register_element_class('style:style', odf_chart_style, family='chart')
