# Import from the Standard Library
from unittest import TestCase, main

from chart_style import odf_create_chart_style, odf_create_chart_title_style
from chart_style import odf_create_chart_legend_style
from chart_style import odf_create_chart_plot_area_style
from chart_style import odf_create_chart_axis_style
from chart_style import odf_create_chart_axis_title_style
from chart_style import odf_create_chart_wall_style

class chart_style_test(TestCase):

    def test_set_auto_postion(self):
        s = odf_create_chart_style("my_name")
        s.set_auto_position("true")
        expected = ('<style:style style:family="chart" style:name="my_name">'
                      '<style:chart-properties chart:auto-position="true"/>'
                    '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_auto_position(self):
        s = odf_create_chart_style("my_name")
        s.set_auto_position("true")
        self.assertEqual(s.get_auto_position(), True)

    def test_set_auto_size(self):
        s = odf_create_chart_style("my_name")
        s.set_auto_size("true")
        expected = ('<style:style style:family="chart" style:name="my_name">'
                      '<style:chart-properties chart:auto-size="true"/>'
                    '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_auto_size(self):
        s = odf_create_chart_style("my_name")
        s.set_auto_size("true")
        self.assertEqual(s.get_auto_size(), True)

    def test_set_font_size(self):
        s = odf_create_chart_style("my_name")
        s.set_font_size("11pt")
        expected = ('<style:style style:family="chart" style:name="my_name">'
                      '<style:text-properties fo:font-size="11pt" '
                        'style:font-size-asian="11pt" '
                        'style:font-size-complex="11pt"/>'
                    '</style:style>')
        self.assertEqual(s.serialize(), expected)
        
    def test_get_font_size(self):
        s = odf_create_chart_style("my_name")
        s.set_font_size("11pt")
        self.assertEqual(s.get_font_size(), "11pt")

    def test_set_stroke(self):
        s = odf_create_chart_style("my_name")
        s.set_stroke("solid", "#b3b3b3")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:graphic-properties draw:stroke="solid" '
                      'svg:stroke-color="#b3b3b3"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_stroke(self):
        s = odf_create_chart_style("my_name")
        s.set_stroke("solid", "#b3b3b3")
        self.assertEqual(s.get_stroke(), "solid")

    def test_get_stroke_color(self):
        s = odf_create_chart_style("my_name")
        s.set_stroke("solid", "#b3b3b3")
        self.assertEqual(s.get_stroke_color(), "#b3b3b3")

    def test_set_fill(self):
        s = odf_create_chart_style("my_name")
        s.set_fill("solid", "#e6e6e6")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:graphic-properties draw:fill="solid" '
                      'draw:fill-color="#e6e6e6"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_fill(self):
        s = odf_create_chart_style("my_name")
        s.set_fill("solid", "#e6e6e6")
        self.assertEqual(s.get_fill(), "solid")

    def test_get_fill_color(self):
        s = odf_create_chart_style("my_name")
        s.set_fill("solid", "#e6e6e6")
        self.assertEqual(s.get_fill_color(), "#e6e6e6")

    def test_set_include_hidden_cells(self):
        s = odf_create_chart_style("my_name")
        s.set_include_hidden_cells("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties '
                      'chart:include-hidden-cells="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_include_hidden_cells(self):
        s = odf_create_chart_style("my_name")
        s.set_include_hidden_cells("true")
        self.assertEqual(s.get_include_hidden_cells(), True)

    def test_set_data_style(self):
        s = odf_create_chart_style("my_name")
        s.set_data_style("st1")
        expected=('<style:style style:family="chart" style:name="my_name" '
                    'style:data-style-name="st1"/>')
        self.assertEqual(s.serialize(), expected)

    def test_get_data_style(self):
        s = odf_create_chart_style("my_name")
        s.set_data_style("st1")
        self.assertEqual(s.get_data_style(), "st1")

    def test_set_display_label(self):
        s = odf_create_chart_style("my_name")
        s.set_display_label("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:display-label="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_display_label(self):
        s = odf_create_chart_style("my_name")
        s.set_display_label("true")
        self.assertEqual(s.get_display_label(), True)

    def test_set_logarithmic(self):
        s = odf_create_chart_style("my_name")
        s.set_logarithmic("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:logarithmic="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_logarithmic(self):
        s = odf_create_chart_style("my_name")
        s.set_logarithmic("true")
        self.assertEqual(s.get_logarithmic(), True)

    def test_set_reverse_direction(self):
        s = odf_create_chart_style("my_name")
        s.set_reverse_direction("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:reverse-direction="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_reverse_direction(self):
        s = odf_create_chart_style("my_name")
        s.set_reverse_direction("true")
        self.assertEqual(s.get_reverse_direction(), True)

    def test_set_line_break(self):
        s = odf_create_chart_style("my_name")
        s.set_line_break("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties text:line-break="true"/>'
                   '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_line_break(self):
        s = odf_create_chart_style("my_name")
        s.set_line_break("true")
        self.assertEqual(s.get_line_break(), True)

    def test_set_axis_position(self):
        s = odf_create_chart_style("my_name")
        s.set_axis_position("0")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:axis-position="0"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_axis_position(self):
        s = odf_create_chart_style("my_name")
        s.set_axis_position("0")
        self.assertEqual(s.get_axis_position(), "0")

    def test_set_rotation_angle(self):
        s = odf_create_chart_style("my_name")
        s.set_rotation_angle("0")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties style:rotation-angle="0"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_rotation_angle(self):
        s = odf_create_chart_style("my_name")
        s.set_rotation_angle("0")
        self.assertEqual(s.get_rotation_angle(), "0")

    def test_set_stroke_width(self):
        s = odf_create_chart_style("my_name")
        s.set_stroke_width("0.1cm")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:graphic-properties svg:stroke-width="0.1cm"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_stroke_width(self):
        s = odf_create_chart_style("my_name")
        s.set_stroke_width("0.1cm")
        self.assertEqual(s.get_stroke_width(), "0.1cm")

    def test_set_symbol_type(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_type("automatic")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:symbol-type="automatic"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_symbol_type(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_type("automatic")
        self.assertEqual(s.get_symbol_type(), "automatic")

    def test_set_series_source(self):
        s = odf_create_chart_style("my_name")
        s.set_series_source("rows")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:series-source="rows"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_series_source(self):
        s = odf_create_chart_style("my_name")
        s.set_series_source("rows")
        self.assertEqual(s.get_series_source(), "rows")

    def test_set_symbol_name(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_name("square")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:symbol-name="square"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_symbol_name(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_name("square")
        self.assertEqual(s.get_symbol_name(), "square")

    def test_set_symbol_width(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_width("0.25cm")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:symbol-width="0.25cm"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_symbol_width(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_width("0.25cm")
        self.assertEqual(s.get_symbol_width(), "0.25cm")

    def test_set_symbol_height(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_height("0.25cm")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:symbol-height="0.25cm"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_symbol_height(self):
        s = odf_create_chart_style("my_name")
        s.set_symbol_height("0.25cm")
        self.assertEqual(s.get_symbol_height(), "0.25cm")

    def test_set_treat_empty_cells(self):
        s = odf_create_chart_style("my_name")
        s.set_treat_empty_cells("leave-gap")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties '
                      'chart:treat-empty-cells="leave-gap"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_treat_empty_cells(self):
        s = odf_create_chart_style("my_name")
        s.set_treat_empty_cells("leave-gap")
        self.assertEqual(s.get_treat_empty_cells(), "leave-gap")

    def test_set_right_angled_axes(self):
        s = odf_create_chart_style("my_name")
        s.set_right_angled_axes("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:right-angled-axes="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_right_angled_axes(self):
        s = odf_create_chart_style("my_name")
        s.set_right_angled_axes("true")
        self.assertEqual(s.get_right_angled_axes(), True)

    def test_set_stacked(self):
        s = odf_create_chart_style("my_name")
        s.set_stacked("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:stacked="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_stacked(self):
        s = odf_create_chart_style("my_name")
        s.set_stacked("true")
        self.assertEqual(s.get_stacked(), True)

    def test_set_vertical(self):
        s = odf_create_chart_style("my_name")
        s.set_vertical("true")
        expected=('<style:style style:family="chart" style:name="my_name">'
                    '<style:chart-properties chart:vertical="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_get_vertical(self):
        s = odf_create_chart_style("my_name")
        s.set_vertical("true")
        self.assertEqual(s.get_vertical(), True)

class quick_chart_style_test(TestCase):

    def test_odf_create_chart_title_style(self):
        s=odf_create_chart_title_style("title", "12pt")
        expected=('<style:style style:family="chart" style:name="title">'
                    '<style:chart-properties chart:auto-position="true"/>'
                    '<style:text-properties fo:font-size="12pt" '
                      'style:font-size-asian="12pt" '
                      'style:font-size-complex="12pt"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_odf_create_chart_legend_style(self):
        s = odf_create_chart_legend_style("legend")
        expected=('<style:style style:family="chart" style:name="legend">'
                    '<style:chart-properties chart:auto-position="true"/>'
                    '<style:graphic-properties draw:stroke="none" '
                      'svg:stroke-color="#b3b3b3" draw:fill="none" '
                      'draw:fill-color="#e6e6e6"/>'
                    '<style:text-properties fo:font-size="8pt" '
                      'style:font-size-asian="8pt" '
                      'style:font-size-complex="8pt"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_odf_create_chart_plot_area_style(self):
        s=odf_create_chart_plot_area_style("plot")
        expected=('<style:style style:family="chart" style:name="plot">'
                    '<style:chart-properties '
                      'chart:include-hidden-cells="false" '
                      'chart:auto-size="true" chart:auto-position="true" '
                      'chart:treat-empty-cells="leave-gap" '
                      'chart:right-angled-axes="true"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_odf_create_chart_axis_style(self):
        s=odf_create_chart_axis_style("axis")
        expected=('<style:style style:family="chart" style:name="axis">'
                    '<style:chart-properties chart:display-label="true" '
                      'chart:logarithmic="false" '
                      'chart:reverse-direction="false" text:line-break="true" '
                      'chart:axis-position="0"/>'
                    '<style:graphic-properties draw:stroke="none" '
                      'svg:stroke-color="#b3b3b3"/>'
                    '<style:text-properties fo:font-size="8pt" '
                      'style:font-size-asian="8pt" '
                      'style:font-size-complex="8pt"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_odf_create_chart_axis_title_style(self):
        s=odf_create_chart_axis_title_style("title", "0")
        expected=('<style:style style:family="chart" style:name="title">'
                    '<style:chart-properties chart:auto-position="true" '
                      'style:rotation-angle="0"/>'
                    '<style:text-properties fo:font-size="9pt" '
                    'style:font-size-asian="9pt" '
                    'style:font-size-complex="9pt"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

    def test_odf_create_chart_wall_style(self):
        s=odf_create_chart_wall_style("wall")
        expected=('<style:style style:family="chart" style:name="wall">'
                    '<style:graphic-properties draw:stroke="solid" '
                    'svg:stroke-color="#b3b3b3" draw:fill="none" '
                    'draw:fill-color="#e6e6e6"/>'
                  '</style:style>')
        self.assertEqual(s.serialize(), expected)

if __name__ == '__main__':
    main()
