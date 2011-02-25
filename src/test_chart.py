import unittest
from chart import odf_create_chart, odf_create_plot_area, create_simple_chart
from chart import add_chart_structure_in_document, attach_chart_to_cell
from chart import divide_range, split_range
from lpod.table import  odf_create_cell
from lpod.document import odf_new_document
from lpod.xmlpart import odf_xmlpart


class TestChart(unittest.TestCase):
    
    def test_odf_create_chart(self):
        chart = odf_create_chart('bar')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm"/>')

        self.assertEqual(chart.serialize(), expected)

    def test_get_class(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_class(), "chart:bar")

    def test_set_class(self):
        chart = odf_create_chart('bar')
        chart.set_class('line')
	expected = ('<chart:chart chart:class="chart:line" svg:width="10cm" '
                      'svg:height="10cm"/>')

        self.assertEqual(chart.serialize(), expected)

    def test_get_size(self):
        chart = odf_create_chart('bar')
        expected = (u'10cm', u'10cm')
        self.assertEqual(chart.get_size(), expected)

    def test_set_size(self):
        chart = odf_create_chart('bar')
        chart.set_size('5cm', '5cm')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="5cm" '
                      'svg:height="5cm"/>')

        self.assertEqual(chart.serialize(), expected)

    def test_get_style(self):
        chart = odf_create_chart('bar')
        chart.set_style('st1')
        expected = "st1"
        self.assertEqual(chart.get_style(), expected)

    def test_set_style(self):
        chart = odf_create_chart('bar')
        chart.set_style('st1')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm" chart:style-name="st1"/>')

        self.assertEqual(chart.serialize(), expected)

    def test_get_title(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_title(), None)
        chart.set_title('title')
        self.assertEqual(chart.get_title(), "title")

    def test_set_title(self):
        chart = odf_create_chart('bar')
        chart.set_title('title')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:title>'
                        '<text:p>title</text:p>'
                      '</chart:title>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)
        
    def test_get_title_style(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_title_style(), None)
        chart.set_title('title')
        chart.set_title_style('st1')
        self.assertEqual(chart.get_title_style(), 'st1')

    def test_set_title_style(self):
        chart = odf_create_chart('bar')
        chart.set_title('title')
        chart.set_title_style('st1')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:title chart:style-name="st1">'
                        '<text:p>title</text:p>'
                      '</chart:title>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_get_title_position(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_title_position(), None)
        chart.set_title('title')
        chart.set_title_position("1cm", "2cm")
        self.assertEqual(chart.get_title_position(), (u'1cm', u'2cm') )

    def test_set_title_position(self):
        chart = odf_create_chart('bar')
        chart.set_title('title')
        chart.set_title_position("1cm", "2cm")
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:title svg:x="1cm" svg:y="2cm">'
                        '<text:p>title</text:p>'
                      '</chart:title>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_get_subtitle(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_subtitle(), None)    
        chart.set_subtitle('subtitle')
        self.assertEqual(chart.get_subtitle(), "subtitle")    
        
    def test_set_subtitle(self):
        chart = odf_create_chart('bar')
        chart.set_subtitle('subtitle')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:subtitle>'
                        '<text:p>subtitle</text:p>'
                      '</chart:subtitle>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_get_subtitle_style(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_subtitle_style(), None)
        chart.set_subtitle('subtitle')
        chart.set_subtitle_style('st1')
        self.assertEqual(chart.get_subtitle_style(), 'st1')

    def test_set_subtitle_style(self):
        chart = odf_create_chart('bar')
        chart.set_subtitle('subtitle')
        chart.set_subtitle_style('st1')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:subtitle chart:style-name="st1">'
                        '<text:p>subtitle</text:p>'
                      '</chart:subtitle>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)
        
    def test_get_subtitle_position(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_subtitle_position(), None)
        chart.set_subtitle('subtitle')
        chart.set_subtitle_position("1cm", "2cm")
        self.assertEqual(chart.get_subtitle_position(), (u'1cm', u'2cm') )

    def test_set_subtitle_position(self):
        chart = odf_create_chart('bar')
        chart.set_subtitle('subtitle')
        chart.set_subtitle_position("1cm", "2cm")
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:subtitle svg:x="1cm" svg:y="2cm">'
                        '<text:p>subtitle</text:p>'
                      '</chart:subtitle>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_get_footer(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_footer(), None)    
        chart.set_footer('footer')
        self.assertEqual(chart.get_footer(), "footer")    
        
    def test_set_footer(self):
        chart = odf_create_chart('bar')
        chart.set_footer('footer')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:footer>'
                        '<text:p>footer</text:p>'
                      '</chart:footer>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_get_footer_style(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_footer_style(), None)
        chart.set_footer('footer')
        chart.set_footer_style('st1')
        self.assertEqual(chart.get_footer_style(), 'st1')

    def test_set_footer_style(self):
        chart = odf_create_chart('bar')
        chart.set_footer('footer')
        chart.set_footer_style('st1')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:footer chart:style-name="st1">'
                        '<text:p>footer</text:p>'
                      '</chart:footer>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)
        
    def test_get_footer_position(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_footer_position(), None)
        chart.set_footer('footer')
        chart.set_footer_position("1cm", "2cm")
        self.assertEqual(chart.get_footer_position(), (u'1cm', u'2cm') )

    def test_set_footer_position(self):
        chart = odf_create_chart('bar')
        chart.set_footer('footer')
        chart.set_footer_position("1cm", "2cm")
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:footer svg:x="1cm" svg:y="2cm">'
                        '<text:p>footer</text:p>'
                      '</chart:footer>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_set_legend(self):
        chart = odf_create_chart('bar')
        chart.set_legend('end')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
                      '<chart:legend chart:legend-position="end"/>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)
        
    def test_get_legend_position(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_legend_position(), None)
        chart.set_legend('end')
        self.assertEqual(chart.get_legend_position(), 'end')
        
    def test_set_legend_alignment(self):
        chart = odf_create_chart('bar')
        chart.set_legend('end')
        chart.set_legend_alignment('center')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
		      '<chart:legend chart:legend-position="end" '
                        'chart:legend-align="center"/>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_get_legend_alignment(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_legend_alignment(), None)
        chart.set_legend('end')
        chart.set_legend_alignment('center')
        self.assertEqual(chart.get_legend_alignment(), 'center')

    def test_set_legend_style(self):
        chart = odf_create_chart('bar')
        chart.set_legend('end')
        chart.set_legend_style('st1')
	expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                      'svg:height="10cm">'
		      '<chart:legend chart:legend-position="end" '
                        'chart:style-name="st1"/>'
                    '</chart:chart>')
        self.assertEqual(chart.serialize(), expected)

    def test_get_legend_style(self):
        chart = odf_create_chart('bar')
        self.assertEqual(chart.get_legend_style(), None)
        chart.set_legend('end')
        chart.set_legend_style('st1')
        self.assertEqual(chart.get_legend_style(), 'st1')

#
#Plot Area
#

class TestPlotArea(unittest.TestCase):

    def test_plot_create_plot_area(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
	expected = ('<chart:plot-area '
                     'table:cell-range-address="Sheet1.A1:Sheet1.A2"/>')

        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_cell_range_address(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        expected = 'Sheet1.A1:Sheet1.A2'
        self.assertEqual(pa.get_cell_range_address(), expected)

    def test_plot_set_cell_range_address(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_cell_range_address('Sheet1.A10:Sheet1.A20')
	expected = ('<chart:plot-area '
                     'table:cell-range-address="Sheet1.A10:Sheet1.A20"/>')

        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_style('st1')
        self.assertEqual(pa.get_style(), "st1")

    def test_plot_set_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_style('st1')
	expected = ('<chart:plot-area '
		     'table:cell-range-address="Sheet1.A1:Sheet1.A2" '
                     'chart:style-name="st1"/>')

        self.assertEqual(pa.serialize(), expected) 

    def test_plot_get_position(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_position("1cm", "2cm")
        self.assertEqual(pa.get_position(), (u'1cm', u'2cm')) 

    def test_plot_set_position(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_position("1cm", "2cm")
	expected = ('<chart:plot-area '
		     'table:cell-range-address="Sheet1.A1:Sheet1.A2" '
                     'svg:x="1cm" svg:y="2cm"/>')

        self.assertEqual(pa.serialize(), expected) 

    def test_plot_get_size(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_size("1cm", "2cm")
        self.assertEqual(pa.get_size(), (u'1cm', u'2cm')) 

    def test_plot_set_size(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_size("1cm", "2cm")
	expected = ('<chart:plot-area '
		     'table:cell-range-address="Sheet1.A1:Sheet1.A2" '
                     'svg:width="1cm" svg:height="2cm"/>')

        self.assertEqual(pa.serialize(), expected) 

    def test_plot_set_labels(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_labels("column")
	expected = ('<chart:plot-area '
                     'table:cell-range-address="Sheet1.A1:Sheet1.A2" '
                     'chart:data-source-has-labels="column"/>')

        self.assertEqual(pa.serialize(), expected)
    
    def test_plot_get_labels(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_labels("column")
        self.assertEqual(pa.get_labels(), "column")

    def test_plot_set_axis(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:axis chart:dimension="x">'
                        '<chart:title><text:p>Axis X</text:p></chart:title>'
                        '<chart:grid chart:class="major"/>'
                      '</chart:axis>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_set_axis_title(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_axis_title('x', "Axis style")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:axis chart:dimension="x">'
                        '<chart:title><text:p>Axis style</text:p></chart:title>'
                        '<chart:grid chart:class="major"/>'
                      '</chart:axis>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_axis_title(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_axis_title('x'), None)
        pa.set_axis('x')
        self.assertEqual(pa.get_axis_title('x'), None)
        pa.set_axis('x', title='Axis X', grid='major')
        self.assertEqual(pa.get_axis_title('x'), "Axis X")

    def test_plot_set_axis_title_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_axis_title('x', "Axis style")
        pa.set_axis_title_style('x', 'st1')
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                        '<chart:axis chart:dimension="x">'
  		          '<chart:title chart:style-name="st1">'
                            '<text:p>Axis style</text:p>'
                          '</chart:title>'
                          '<chart:grid chart:class="major"/>'
                        '</chart:axis>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_axis_title_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_axis_title('x', "Axis style")
        pa.set_axis_title_style('x', 'st1')
        self.assertEqual(pa.get_axis_title_style('x'), "st1")

    def test_plot_set_axis_grid(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_axis_grid('x', "minor")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:axis chart:dimension="x">'
                        '<chart:title><text:p>Axis X</text:p></chart:title>'
                        '<chart:grid chart:class="minor"/>'
                      '</chart:axis>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)
    
    def test_plot_get_axis_grid(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_axis_grid('x'), None)
        pa.set_axis('x')
        self.assertEqual(pa.get_axis_grid('x'), None)
        pa.set_axis('x', title='Axis X', grid='major')
        self.assertEqual(pa.get_axis_grid('x'), "major")

    def test_plot_get_axis_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_axis_style('x'), None)
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_axis_style('x', 'st1')
        self.assertEqual(pa.get_axis_style('x'), "st1")

    def test_plot_set_axis_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_axis_style('x', 'st1')
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:axis chart:dimension="x" chart:style-name="st1">'
                        '<chart:title><text:p>Axis X</text:p></chart:title>'
                        '<chart:grid chart:class="major"/>'
                      '</chart:axis>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_grid_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_grid_style('x'), None)
        pa.set_axis('x', title='Axis X', grid='major')
        self.assertEqual(pa.get_grid_style('x'), None)
        pa.set_grid_style('x', 'st1')
        self.assertEqual(pa.get_grid_style('x'), "st1")

    def test_plot_set_grid_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_grid_style('x', 'st1')
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:axis chart:dimension="x">'
                        '<chart:title><text:p>Axis X</text:p></chart:title>'
			'<chart:grid chart:class="major" '
                          'chart:style-name="st1"/>'
                      '</chart:axis>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_set_categories(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x')
        pa.set_categories('x', "Sheet1.A1:Sheet1.A2")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:axis chart:dimension="x">'
		        '<chart:categories '
                        'table:cell-range-address="Sheet1.A1:Sheet1.A2"/>'
                      '</chart:axis>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_categories(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_axis('x', title='Axis X', grid='major')
        pa.set_categories('x', "Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_categories('x'), "Sheet1.A1:Sheet1.A2")


    def test_plot_set_chart_series(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
		      '<chart:series '
                        'chart:values-cell-range-address="Sheet1.A1:Sheet1.A2" '
                        'chart:class="chart:line"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_chart_series_values(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_chart_series_values(), [])
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
        self.assertEqual(pa.get_chart_series_values(), [u'Sheet1.A1:Sheet1.A2'])

    def test_plot_set_chart_series_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
        pa.set_chart_series_style("Sheet1.A1:Sheet1.A2", "st1")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
		      '<chart:series '
		      'chart:values-cell-range-address="Sheet1.A1:Sheet1.A2" '
                      'chart:class="chart:line" chart:style-name="st1"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_chart_series_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_chart_series_style("Sheet1.A1:Sheet1.A2"), None)
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
        pa.set_chart_series_style("Sheet1.A1:Sheet1.A2", "st1")
	self.assertEqual(pa.get_chart_series_style("Sheet1.A1:Sheet1.A2"),
                                                                      "st1")

    def test_plot_set_chart_series_label(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
        pa.set_chart_series_label("Sheet1.A1:Sheet1.A2", "Sheet1.A1:Sheet1.A2")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
		      '<chart:series '
		        'chart:values-cell-range-address="Sheet1.A1:Sheet1.A2" '
			'chart:class="chart:line" '
                        'chart:label-cell-address="Sheet1.A1:Sheet1.A2"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_chart_series_label(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_chart_series_label("Sheet1.A1:Sheet1.A2"), None)
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
        pa.set_chart_series_label("Sheet1.A1:Sheet1.A2", "Sheet1.A1:Sheet1.A2")
	self.assertEqual(pa.get_chart_series_label("Sheet1.A1:Sheet1.A2"),
                                                         "Sheet1.A1:Sheet1.A2")

    def test_plot_set_chart_series_domain(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
        pa.set_chart_series_domain("Sheet1.A1:Sheet1.A2", "Sheet1.A1:Sheet1.A2")
	expected = ('<chart:plot-area '
                     'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                     '<chart:series '
                       'chart:values-cell-range-address="Sheet1.A1:Sheet1.A2" '
                       'chart:class="chart:line">'
                       '<chart:domain '
                       'table:cell-range-address="Sheet1.A1:Sheet1.A2"/>'
                     '</chart:series>'
                   '</chart:plot-area>')

        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_chart_series_domain(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_chart_series_domain("Sheet1.A1:Sheet1.A2"), None)
        pa.set_chart_series("Sheet1.A1:Sheet1.A2", "line")
        pa.set_chart_series_domain("Sheet1.A1:Sheet1.A2", "Sheet1.A1:Sheet1.A2")
	self.assertEqual(pa.get_chart_series_domain("Sheet1.A1:Sheet1.A2"),
                                                          "Sheet1.A1:Sheet1.A2")

    def test_plot_set_floor(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_floor(width="8cm", style="st1")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:floor svg:width="8cm" chart:style-name="st1"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_floor_width(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_floor_width(), None)
        pa.set_floor(width="8cm")
        self.assertEqual(pa.get_floor_width(), "8cm")
      
    def test_plot_get_floor_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_floor_style(), None)
        pa.set_floor(style="st1")
        self.assertEqual(pa.get_floor_style(), "st1") 

    def test_plot_set_wall(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_wall(width="8cm", style="st1")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:wall svg:width="8cm" chart:style-name="st1"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_wall_width(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_wall_width(), None)
        pa.set_wall(width="8cm")
        self.assertEqual(pa.get_wall_width(), "8cm")
      
    def test_plot_get_wall_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_wall_style(), None)
        pa.set_wall(style="st1")
        self.assertEqual(pa.get_wall_style(), "st1")

    def test_plot_set_stock_loss_marker(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_stock_loss_marker(style="st1")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:stock-loss-marker chart:style-name="st1"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_stock_loss_marker_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_stock_loss_marker_style(), None)
        pa.set_stock_loss_marker(style="st1")
        self.assertEqual(pa.get_stock_loss_marker_style(), "st1")

    def test_plot_set_stock_gain_marker(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_stock_gain_marker(style="st1")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:stock-gain-marker chart:style-name="st1"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_stock_gain_marker_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_stock_gain_marker_style(), None)
        pa.set_stock_gain_marker(style="st1")
        self.assertEqual(pa.get_stock_gain_marker_style(), "st1")

    def test_plot_set_stock_range_line(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        pa.set_stock_range_line(style="st1")
	expected = ('<chart:plot-area '
                      'table:cell-range-address="Sheet1.A1:Sheet1.A2">'
                      '<chart:stock-range-line chart:style-name="st1"/>'
                    '</chart:plot-area>')
        self.assertEqual(pa.serialize(), expected)

    def test_plot_get_stock_range_line_style(self):
        pa = odf_create_plot_area("Sheet1.A1:Sheet1.A2")
        self.assertEqual(pa.get_stock_range_line_style(), None)
        pa.set_stock_range_line(style="st1")
        self.assertEqual(pa.get_stock_range_line_style(), "st1")





class TestQuickCharts(unittest.TestCase):
    def test_create_simple_chart(self):
        chart = create_simple_chart('bar', 'Sheet1.A1:Sheet1.B2', 'title' )
        expected = ('<chart:chart chart:class="chart:bar" svg:width="10cm" '
                    'svg:height="10cm">'
                    '<chart:title><text:p>title</text:p></chart:title>'
                    '<chart:legend chart:legend-position="bottom"/>'
                    '<chart:plot-area '
                       'table:cell-range-address="Sheet1.A1:Sheet1.B2">'
                     '<chart:axis chart:dimension="x">'
                       '<chart:categories table:cell-range-address=""/>'
                     '</chart:axis>'
                     '<chart:axis chart:dimension="y">'
                       '<chart:grid chart:class="minor"/>'
                     '</chart:axis>'
                     '<chart:series '
                       'chart:values-cell-range-address="Sheet1.A1:Sheet1.A2" '
                       'chart:class="chart:bar"/>'
                     '<chart:series '
                       'chart:values-cell-range-address="Sheet1.B1:Sheet1.B2" '
                       'chart:class="chart:bar"/>'
                     '<chart:floor/>'
                     '<chart:wall/>'
                    '</chart:plot-area>'
                   '</chart:chart>')

        self.assertEqual(chart.serialize(), expected)

class UsefulFunctions(unittest.TestCase):
    def test_split_range(self):
	expected = {'labels': 'sheet.B1:sheet.C1', 'data': 'sheet.B2:sheet.C4',
                                                 'legend': 'sheet.A2:sheet.A4'}

	self.assertEqual(split_range("sheet.A1:sheet.C4", legend="column",
                                                     x_labels="row"), expected)

	expected = {'labels': 'sheet.A1:sheet.C1', 'data': 'sheet.A2:sheet.C4',
                                                                  'legend': ''}

	self.assertEqual(split_range("sheet.A1:sheet.C4", legend="none",
                                                     x_labels="row"), expected)

	expected = {'labels': '', 'data': 'sheet.B1:sheet.C4', 'legend':
                                                           'sheet.A1:sheet.A4'}

	self.assertEqual(split_range("sheet.A1:sheet.C4", legend="column",
                                                    x_labels="none"), expected)


    def test_divide_range(self):
	expected = ['sheet.A1:sheet.A3', 'sheet.B1:sheet.B3',
                                                         'sheet.C1:sheet.C3']

        self.assertEqual(divide_range("sheet.A1:sheet.C3"), expected)
	expected = ['sheet.A1:sheet.C1', 'sheet.A2:sheet.C2',
                                                         'sheet.A3:sheet.C3']

        self.assertEqual(divide_range("sheet.A1:sheet.C3", 'rows'), expected)
	self.assertRaises(AttributeError, divide_range, "sheet.A1:sheet.C3",
                                                                      "test")


    def test_add_chart_structure_in_document(self):
        doc = odf_new_document('spreadsheet')
        name = add_chart_structure_in_document(doc)
        chart_content = odf_xmlpart(name+'/content.xml', doc)
        expected =(
          '<?xml version="1.0" encoding="UTF-8"?>\n'
          '<office:document-content '
          'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
          'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
          'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
          'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
          'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" '
          'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
          'xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" '
          'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
          'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" '
          'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
          'xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" '
          'xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" '
          'xmlns:math="http://www.w3.org/1998/Math/MathML" '
          'xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" '
          'xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" '
          'xmlns:ooo="http://openoffice.org/2004/office" '
          'xmlns:ooow="http://openoffice.org/2004/writer" '
          'xmlns:oooc="http://openoffice.org/2004/calc" '
          'xmlns:dom="http://www.w3.org/2001/xml-events" '
          'xmlns:xforms="http://www.w3.org/2002/xforms" '
          'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
          'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
          'xmlns:rpt="http://openoffice.org/2005/report" '
          'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" '
          'xmlns:xhtml="http://www.w3.org/1999/xhtml" '
          'xmlns:grddl="http://www.w3.org/2003/g/data-view#" '
          'xmlns:tableooo="http://openoffice.org/2009/table" '
          'xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" '
          'xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" '
          'office:version="1.2" '
          'grddl:transformation="http://docs.oasis-open.org/office/1.2/xslt/odf2rdf.xsl">\n\n  '
          '<office:automatic-styles/>\n  '
          '<office:body>\n    '
          '<office:chart/>\n  '
          '</office:body>\n'
          '</office:document-content>')
        self.assertEqual(chart_content.serialize(), expected)

    def test_attach_chart_to_cell(self):
        cell = odf_create_cell()
        cell = attach_chart_to_cell('Object', cell)
        expected=('<table:table-cell>'
                    '<draw:frame svg:width="10cm" svg:height="10cm" '
                      'draw:z-index="0">'
                      '<draw:object xlink:href="./Object" xlink:type="simple" '
                        'xlink:show="embed" xlink:actuate="onLoad"/>'
                    '</draw:frame>'
                  '</table:table-cell>')
        self.assertEqual(cell.serialize(), expected)
        

if __name__ == '__main__':
    unittest.main()
