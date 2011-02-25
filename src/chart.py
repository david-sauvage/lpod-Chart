from lpod.element import odf_create_element, register_element_class, odf_element
from lpod.frame import odf_create_frame
from lpod.element import odf_create_element
from lpod.container import odf_new_container
from lpod.table import _get_cell_coordinates, _digit_to_alpha
from lpod.utils import _get_abspath
from string import splitfields
from os import getcwd
from sys import path
from os.path import exists
    
#
#Chart
#

def odf_create_chart(chart_class, size=('10cm', '10cm'), title=None,
				subtitle=None, legend_position=None):

    """
    Create a chart structure for basic use
    
    Arguments
         class -- 'line' 'area' 'circle' 'ring' 'scatter' 'radar'
                  'bar' 'stock' 'bubble' 'surface' 'gant'
         size -- (str, str)
         title -- str
         subtitle -- str
         legend_position -- None 'start' 'end' 'top' 'bottom' 
                            'top-start' 'bottom-start'
                            'top-end' 'bottom-end'
    """
    element = odf_create_element('chart:chart')
    element.set_class(chart_class)
    element.set_size(size[0], size[1])
    if title is not None:
        element.set_title(title)
    if subtitle is not None:
        element.set_subtitle(subtitle) 
    if legend_position is not None:
        element.set_legend(legend_position)
        
    return element


class odf_chart(odf_element):
    """ This class implements an xml element that represents a chart"""
#class
    def get_class(self):
        return self.get_attribute('chart:class')

    def set_class(self, chart_class):
        """
        chart_class -- 'line' 'area' 'circle' 'ring' 'scatter' 'radar'
                       'bar' 'stock' 'bubble' 'surface' 'gant'
        """          
        self.set_attribute('chart:class', "chart:"+chart_class)

#size
    def get_size(self):
        h = self.get_attribute('svg:height')
        w = self.get_attribute('svg:width')
        return (w, h)

    def set_size(self, w, h):
        """ arguments -- str """
        self.set_attribute('svg:width', w)
        self.set_attribute('svg:height', h)

#style
    def get_style(self):
        return self.get_attribute('chart:style-name')

    def set_style(self, style):
        """ style -- str """
        self.set_attribute('chart:style-name', style)

#title
    def get_title(self):
        elt_title = self.get_element('chart:title')
        if elt_title is not None:
            return elt_title.get_text_content()
        else:
            return None

    def set_title(self, text):
        """ text -- str """
        elt_title = self.get_element('chart:title')
        if elt_title is None:
            elt_title = odf_create_element('chart:title')
            self.append(elt_title)
        elt_title.set_text_content(text)
      
    def get_title_style(self):
        elt_title = self.get_element('chart:title')
        if elt_title is not None:
            return elt_title.get_attribute('chart:style-name')
        else:
            return None

    def set_title_style(self, style):
        """ style -- str """
        elt_title = self.get_element('chart:title')
        if elt_title is None:
            elt_title = odf_create_element('chart:title')
            self.append(elt_title)
        elt_title.set_attribute('chart:style-name', style)

    def get_title_position(self):
        elt_title = self.get_element('chart:title')
        if elt_title is not None:
            x = elt_title.get_attribute('svg:x')
            y = elt_title.get_attribute('svg:y')
            return (x, y)
        else:
            return None
    
    def set_title_position(self, x, y):
        """ arguments -- str """
        elt_title = self.get_element('chart:title')
        if elt_title is None:
            elt_title = odf_create_element('chart:title')
            self.append(elt_title)
        elt_title.set_attribute('svg:x', x)
        elt_title.set_attribute('svg:y', y)

#subtitle
    def get_subtitle(self):
        elt_subtitle = self.get_element('chart:subtitle')
        if elt_subtitle is not None:
            return elt_subtitle.get_text_content()
        else:
            return None

    def set_subtitle(self, text):
        """ text -- str """
        elt_subtitle = self.get_element('chart:subtitle')
        if elt_subtitle is None:
            elt_subtitle = odf_create_element('chart:subtitle')
            self.append(elt_subtitle)
        elt_subtitle.set_text_content(text)

    def get_subtitle_style(self):
        elt_subtitle = self.get_element('chart:subtitle')
        if elt_subtitle is not None:
            return elt_subtitle.get_attribute('chart:style-name')
        else:
            return None
    
    def set_subtitle_style(self, style):
        """ style -- str """
        elt_subtitle = self.get_element('chart:subtitle')
        if elt_subtitle is None:
            elt_subtitle = odf_create_element('chart:subtitle')
            self.append(elt_subtitle)
        elt_subtitle.set_attribute('chart:style-name', style)

    def get_subtitle_position(self):
        elt_subtitle = self.get_element('chart:subtitle')
        if elt_subtitle is not None:
            x=elt_subtitle.get_attribute('svg:x')
            y=elt_subtitle.get_attribute('svg:y')
            return (x, y)
        else:
            return None
    
    def set_subtitle_position(self, x, y):
        """ arguments -- str """
        elt_subtitle = self.get_element('chart:subtitle')
        if elt_subtitle is None:
            elt_subtitle = odf_create_element('chart:subtitle')
            self.append(elt_subtitle)
        elt_subtitle.set_attribute('svg:x', x)
        elt_subtitle.set_attribute('svg:y', y)


#footer
    def get_footer(self):
        elt_footer = self.get_element('chart:footer')
        if elt_footer is not None:
            return elt_footer.get_text_content()
        else:
            return None

    def set_footer(self, text):
        """ test -- str """
        elt_footer = self.get_element('chart:footer')
        if elt_footer is None:
            elt_footer = odf_create_element('chart:footer')
            self.append(elt_footer)
        elt_footer.set_text_content(text)

    def get_footer_style(self):
        elt_footer = self.get_element('chart:footer')
        if elt_footer is not None:
            return elt_footer.get_attribute('chart:style-name')
        else:
            return None
    
    def set_footer_style(self, style):
        """ style -- str """
        elt_footer = self.get_element('chart:footer')
        if elt_footer is None:
            elt_footer = odf_create_element('chart:footer')
            self.append(elt_footer)
        elt_footer.set_attribute('chart:style-name', style)

    def get_footer_position(self):
        elt_footer = self.get_element('chart:footer')
        if elt_footer is not None:
            x=elt_footer.get_attribute('svg:x')
            y=elt_footer.get_attribute('svg:y')
            return (x, y)
        else:
            return None
    
    def set_footer_position(self, x, y):
        """ arguments -- str """
        elt_footer = self.get_element('chart:footer')
        if elt_footer is None:
            elt_footer = odf_create_element('chart:footer')
            self.append(elt_footer)
        elt_footer.set_attribute('svg:x', x)
        elt_footer.set_attribute('svg:y', y)



#legend
    def set_legend(self, position):
        """
        position -- 'start' 'end' 'top' 'bottom' 'top-start' 'bottom-start'
                    'top-end' 'bottom-end'
        """
        elt_legend = self.get_element('chart:legend')
        if elt_legend is None:
            elt_legend = odf_create_element('chart:legend')
            self.append(elt_legend)
        elt_legend.set_attribute('chart:legend-position', position)

    def get_legend_position(self):
        elt_legend = self.get_element('chart:legend')
        if elt_legend is not None:
            return elt_legend.get_attribute('chart:legend-position')
        else:
            return None

    def set_legend_alignment(self, align):
        """ align -- 'start' 'center' 'end' """
        elt_legend = self.get_element('chart:legend')
        if elt_legend is None:
            elt_legend = odf_create_element('chart:legend')
            self.append(elt_legend)
        elt_legend.set_attribute('chart:legend-align', align)

    def get_legend_alignment(self):
        elt_legend = self.get_element('chart:legend')
        if elt_legend is not None:
            return elt_legend.get_attribute('chart:legend-align')
        else:
            return None

    def set_legend_style(self, style):
        """ style -- str """
        elt_legend = self.get_element('chart:legend')
        if elt_legend is None:
            elt_legend = odf_create_element('chart:legend')
            self.append(elt_legend)
        elt_legend.set_attribute('chart:style-name', style)

    def get_legend_style(self):
        elt_legend = self.get_element('chart:legend')
        if elt_legend is not None:
            return elt_legend.get_attribute('chart:style-name')
        else:
            return None

#
#Plot Area
#

def odf_create_plot_area(cell_range):
    element = odf_create_element('chart:plot-area')
    element.set_cell_range_address(cell_range)
    return element

class odf_plot_area(odf_element):
    """
    This class implements an xml element for a plot area used
    to display the chart
    """

    def get_cell_range_address(self):
        return self.get_attribute('table:cell-range-address')

    def set_cell_range_address(self, range_address):
        """ range_address -- str like "Sheet1.A1:Sheet1.A10" """
        self.set_attribute('table:cell-range-address', range_address)

    def get_style(self):
        return self.get_attribute('chart:style-name')

    def set_style(self, style):
        """ style -- str """
        self.set_attribute('chart:style-name', style)

    def get_position(self):
        x = self.get_attribute('svg:x')
        y = self.get_attribute('svg:y')
        return (x, y)

    def set_position(self, x, y):
        """ arguments -- str """
        self.set_attribute('svg:x', x)
        self.set_attribute('svg:y', y)

    def get_size(self):
        w = self.get_attribute('svg:width')
        h = self.get_attribute('svg:height')
        return (w, h)

    def set_size(self, w, h):
        """ arguments -- str """
        self.set_attribute('svg:width', w)
        self.set_attribute('svg:height', h)

    def set_labels(self, value):
        """ value -- 'none' 'column' 'row' 'both' """
        self.set_attribute('chart:data-source-has-labels', value)

    def get_labels(self):
        return self.get_attribute("chart:data-source-has-labels")

#Axis
#It might have three axis elements so they are defined by their dimension
    def set_axis(self, dimension, title=None, grid=None):
        """
        dimension -- 'x' 'y' 'z'
        title -- str
        grid -- None 'major' 'minor'
        """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is None:
            axis = odf_create_element('chart:axis')
            axis.set_attribute("chart:dimension", dimension)
            self.append(axis)

        if title is not None:
           self.set_axis_title(dimension, title)

        if grid is not None:
            self.set_axis_grid(dimension, grid)


    def set_axis_title(self, dimension, title):
        """
        dimension -- 'x' 'y' 'z'
        title -- str
        """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is None:
            self.set_axis(dimension)
	    axis = self.get_element("//chart:axis[@chart:dimension='" +
                                                              dimension + "']")

        axis_title = axis.get_element('chart:title')
        if axis_title is None:
            axis_title = odf_create_element('chart:title')
            axis.append(axis_title)
        axis_title.set_text_content(title)   

    def get_axis_title(self, dimension):
        """ dimension -- 'x' 'y' 'z' """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            title = axis.get_element('chart:title')
            if title is not None:
                return title.get_text_content()
            else:
                return None
        else:
            return None

    def set_axis_title_style(self, dimension, style):
        """ dimension -- 'x' 'y' 'z' """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is None:
            self.set_axis(dimension)
	    axis = self.get_element("//chart:axis[@chart:dimension='" +
                                                              dimension + "']")

        axis_title = axis.get_element('chart:title')
        if axis_title is None:
            axis_title = odf_create_element('chart:title')
            axis.append(axis_title)
        axis_title.set_attribute("chart:style-name", style)

    def get_axis_title_style(self, dimension):
        """ dimension -- 'x' 'y' 'z' """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            title = axis.get_element('chart:title')
            if title is not None:
                return title.get_attribute("chart:style-name")
            else:
                return None
        else:
            return None

    def set_axis_grid(self, dimension, grid):
        """
        dimension -- 'x' 'y' 'z'
        grid -- 'major' 'minor'
        """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        axis_grid = axis.get_element('chart:grid')
        if axis_grid is None:
            axis_grid = odf_create_element('chart:grid')
            axis.append(axis_grid)
        axis_grid.set_attribute('chart:class', grid)

    def get_axis_grid(self, dimension):
        """  dimension -- 'x' 'y' 'z' """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            grid = axis.get_element('chart:grid')
            if grid is not None:
                return grid.get_attribute('chart:class')
            else:
                return None
        else:
            return None

    def get_axis_style(self, dimension):
        """ dimension -- 'x' 'y' 'z' """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            return axis.get_attribute("chart:style-name")
        else:
            return None

    def set_axis_style(self, dimension, style):
        """
        style -- str
        dimension -- 'x' 'y' 'z'
        """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            axis.set_attribute("chart:style-name", style)
        else:
            raise "axis element not found"

    def get_grid_style(self, dimension):
        """ dimension -- 'x' 'y' 'z' """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            grid = axis.get_element('chart:grid')
            if grid is not None:
                return grid.get_attribute("chart:style-name")
            else:
                return None
        else:
            return None

    def set_grid_style(self, dimension, style):
        """
        style -- str
        dimension -- 'x' 'y' 'z'
        """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            grid = axis.get_element('chart:grid')
            if grid is not None:
                grid.set_attribute("chart:style-name", style)
            else:
                raise "grid element not found"
        else:
            raise "axis element not found"
        
    def set_categories(self, dimension, range_address):
        """
        dimension -- 'x' 'y' 'z'
        range_address -- str
        """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            cat = axis.get_element("chart:categories")
            if cat is None:
                cat = odf_create_element("chart:categories")
                axis.append(cat)
            cat.set_attribute("table:cell-range-address", range_address)
        else:
            raise "axis element not found"

    def get_categories(self, dimension):
        """ dimension -- 'x' 'y' 'z' """
	axis = self.get_element("//chart:axis[@chart:dimension='" + dimension +
                                                                          "']")

        if axis is not None:
            cat = axis.get_element("chart:categories")
            if cat is not None:
                return cat.get_attribute("table:cell-range-address")
            else:
                return None
        else:
            return None
#Series
#A plot area might have many chart:series element
#chart:series are defined with their values attribute        
    def set_chart_series(self, values, chart_class):
        """
        values -- str like "Sheet1.A1:Sheet1.A2"
        chart_class -- 'line' 'area' 'circle' 'ring' 'scatter' 'radar'
                       'bar' 'stock' 'bubble' 'surface' 'gant'
        """
	series = \
	    self.get_element("chart:series[@chart:values-cell-range-address='"
                                                              + values + "']")

        if series is None:
            series = odf_create_element('chart:series')
            series.set_attribute('chart:values-cell-range-address', values) 
            self.append(series)
        series.set_attribute('chart:class', "chart:"+chart_class) 

    def get_chart_series_values(self):
        series = self.get_elements("//chart:series")
	return [s.get_attribute("chart:values-cell-range-address") for s in
                                                                       series]


    def set_chart_series_style(self, values, style):
        """ style -- str """
	series = \
	     self.get_element("chart:series[@chart:values-cell-range-address='"
                                                               + values + "']")

        if series is not None:
            series.set_attribute("chart:style-name", style)
        else:
            raise "Series Element not found"

    def get_chart_series_style(self, values):
	series = \
	      self.get_element("chart:series[@chart:values-cell-range-address='"
                                                                + values + "']")

        if series is not None:
            return series.get_attribute("chart:style-name")
        else:
            return None

    def set_chart_series_label(self, values, label):
        """ label -- str """
	series = \
	      self.get_element("chart:series[@chart:values-cell-range-address='"
                                                                + values + "']")

        if series is not None:
            series.set_attribute("chart:label-cell-address", label)
        else:
            raise "Series Element not found"

    def get_chart_series_label(self, values):
	series = \
	      self.get_element("chart:series[@chart:values-cell-range-address='"
                                                                + values + "']")

        if series is not None:
            return series.get_attribute("chart:label-cell-address")
        else:
            return None
        
    def set_chart_series_domain(self, values, cell_range):
        """ cell_range -- str like 'Sheet1.A1:Sheet.A10' """
	series = \
	      self.get_element("chart:series[@chart:values-cell-range-address='"
                                                                + values + "']")

        if series is not None:
            domain = odf_create_element('chart:domain')
            domain.set_attribute("table:cell-range-address", cell_range)
            series.append(domain)
        else:
            raise "Series Element not found"

    def get_chart_series_domain(self, values):
	domain = \
	      self.get_element("chart:series[@chart:values-cell-range-address='"
                                                   + values + "']/chart:domain")

        if domain is not None:
            return domain.get_attribute("table:cell-range-address")
        else:
            return None

#Floor
    def set_floor(self, width=None, style=None):
        """ arguments -- str """
        floor = self.get_element("chart:floor")
        if floor is None:
            floor = odf_create_element('chart:floor')
            self.append(floor)

        if width is not None:
            floor.set_attribute("svg:width", width)

        if style is not None:
            floor.set_attribute("chart:style-name", style)

    def get_floor_width(self):
        floor = self.get_element("chart:floor")
        if floor is not None:
            return floor.get_attribute("svg:width")
        else:
            return None
        
    def get_floor_style(self):
        floor = self.get_element("chart:floor")
        if floor is not None:
            return floor.get_attribute("chart:style-name")
        else:
            return None

#Wall
    def set_wall(self, width=None, style=None):
        """ arguments -- str """
        wall = self.get_element("chart:wall")
        if wall is None:
            wall = odf_create_element('chart:wall')
            self.append(wall)

        if width is not None:
            wall.set_attribute("svg:width", width)

        if style is not None:
            wall.set_attribute("chart:style-name", style)

    def get_wall_width(self):
        wall = self.get_element("chart:wall")
        if wall is not None:
            return wall.get_attribute("svg:width")
        else:
            return None
        
    def get_wall_style(self):
        wall = self.get_element("chart:wall")
        if wall is not None:
            return wall.get_attribute("chart:style-name")
        else:
            return None

#stock

    def set_stock_loss_marker(self, style=None):
        elt = self.get_element("chart:stock-loss-marker")
        if elt is None:
            elt = odf_create_element("chart:stock-loss-marker")
            self.append(elt)
        if style is not None:
            elt.set_attribute("chart:style-name", style)

    def get_stock_loss_marker_style(self):
        elt = self.get_element("chart:stock-loss-marker")
        if elt is not None:
            return elt.get_attribute("chart:style-name")
        else:
            return None

    def set_stock_gain_marker(self, style=None):
        elt = self.get_element("chart:stock-gain-marker")
        if elt is None:
            elt = odf_create_element("chart:stock-gain-marker")
            self.append(elt)
        if style is not None:
            elt.set_attribute("chart:style-name", style)

    def get_stock_gain_marker_style(self):
        elt = self.get_element("chart:stock-gain-marker")
        if elt is not None:
            return elt.get_attribute("chart:style-name")
        else:
            return None

    def set_stock_range_line(self, style=None):
        elt = self.get_element("chart:stock-range-line")
        if elt is None:
            elt = odf_create_element("chart:stock-range-line")
            self.append(elt)
        if style is not None:
            elt.set_attribute("chart:style-name", style)

    def get_stock_range_line_style(self):
        elt = self.get_element("chart:stock-range-line")
        if elt is not None:
            return elt.get_attribute("chart:style-name")
        else:
            return None

###
#Quick charts
###

def create_simple_chart(chart_type, cell_range, chart_title=None,
		data_in_columns=True, legend="none", x_labels="none"):
    """
    Create a complete chart element for basics chart creation
    Legend and x_labels allows user to inform that the first row
    (or column) contain legend (or x_labels) data       

    chart_type -- 'line' 'area' 'circle' 'ring' 'scatter' 'radar'
                  'bar' 'stock' 'bubble' 'surface' 'gant'
    cell_range -- str like 'Sheet.A1:Sheet.D5'
    chart_title -- str
    data_in_columns -- boolean 
    legend, x_labels -- 'none' 'column' 'row'
    """
    
    chart = odf_create_chart(chart_type, title=chart_title,
                                                      legend_position='bottom')

    po = odf_create_plot_area(cell_range)
    po.set_axis('x')
    po.set_axis('y', grid='minor')

    datas = split_range(cell_range, legend, x_labels)
    po.set_categories('x', datas["labels"])

    legend_list = ()
    if legend == "column":
        legend_list = divide_range(datas["legend"], by="rows")
    elif legend == "row":
        legend_list = divide_range(datas["legend"], by="columns")
        
    i = 0
    if data_in_columns:
        range_list = divide_range(datas["data"])
        for r in range_list:
            po.set_chart_series(r, chart_type)
            if legend is not "none":
                po.set_chart_series_label(r, legend_list[i])
                i = i+1
            
    else:
        range_list = divide_range(datas["data"], by="rows")
        for r in range_list:
            po.set_chart_series(r, chart_type)
            if legend is not "none":
                po.set_chart_series_label(r, legend_list[i])
                i = i+1

    po.set_floor()
    po.set_wall()
    chart.append(po)

    return chart

###
#Useful functions
###
def split_range(cell_range, legend="row", x_labels="column"):
    """
    returns a dictionnary with data_range, legend_range and labels for the
    x-axis

    cell_range - str like Sheet.A1:SheetD5
    legend, x_labels - 'none', 'row', 'column'
    """

    my_dict = {"data":cell_range, "legend":"", "labels":""}
    sheet1=splitfields(splitfields(cell_range, ":")[0],".")[0]
    cell1=splitfields(splitfields(cell_range, ":")[0],".")[1]
    sheet2=splitfields(splitfields(cell_range, ":")[1],".")[0]
    cell2=splitfields(splitfields(cell_range, ":")[1],".")[1]
    
    tmp_coord1 = _get_cell_coordinates(cell1)
    tmp_coord2 = _get_cell_coordinates(cell2)

    #we need the coordinate in a crescent order
    coord1 = (min(tmp_coord1[0], tmp_coord2[0]),min(tmp_coord1[1],
                                                                tmp_coord2[1]))

    coord2 = (max(tmp_coord1[0], tmp_coord2[0]),max(tmp_coord1[1],
                                                                tmp_coord2[1]))


    if legend is not 'none' and x_labels is not 'none':
        #if we have both conditions, we have to delete the first cell
        if legend == "row":
	    my_dict["legend"] = sheet1 + "." + _digit_to_alpha(coord1[0]+1) + \
			        str(coord1[1]+1) + ":"+sheet2 + "." + \
                                _digit_to_alpha(coord2[0]) + str(coord1[1]+1)

        elif legend == "column":
	    my_dict["legend"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
				str(coord1[1]+2) + ":"+sheet2 + "." + \
                                _digit_to_alpha(coord1[0]) + str(coord2[1]+1) 

        if x_labels == "row":
	    my_dict["labels"] = sheet1 + "." + _digit_to_alpha(coord1[0]+1) + \
				str(coord1[1]+1) + ":"+sheet2 + "." + \
                                _digit_to_alpha(coord2[0]) + str(coord1[1]+1)

        elif x_labels == "column":
	    my_dict["labels"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
				str(coord1[1]+2) + ":"+sheet2 + "." + \
                                _digit_to_alpha(coord1[0]) + str(coord2[1]+1)


	my_dict["data"] = sheet1 + "." + _digit_to_alpha(coord1[0]+1) + \
			  str(coord1[1]+2) + ":"+sheet2 + "." + \
                          _digit_to_alpha(coord2[0]) + str(coord2[1]+1)

    else:
        if legend == "row":
	    my_dict["legend"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
				str(coord1[1]+1) + ":"+sheet2 + "." + \
                                _digit_to_alpha(coord2[0]) + str(coord1[1]+1)

	    my_dict["data"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
			      str(coord1[1]+2) + ":" + sheet2 + "." + \
                              _digit_to_alpha(coord2[0]) + str(coord2[1]+1)

        elif legend == "column":
	    my_dict["legend"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
				str(coord1[1]+1) + ":"+sheet2 + "." + \
                                _digit_to_alpha(coord1[0]) + str(coord2[1]+1)

	    my_dict["data"] = sheet1 + "." + _digit_to_alpha(coord1[0]+1) + \
			      str(coord1[1]+1) + ":"+sheet2 + \
                              "."+_digit_to_alpha(coord2[0]) + str(coord2[1]+1)

        if x_labels == "row":
	    my_dict["labels"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
				str(coord1[1]+1) + ":" + sheet2 + "." + \
                                _digit_to_alpha(coord2[0]) + str(coord1[1]+1)

	    my_dict["data"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
			      str(coord1[1]+2) + ":"+sheet2 + "." + \
                              _digit_to_alpha(coord2[0]) + str(coord2[1]+1)

        elif x_labels == "column":
	    my_dict["labels"] = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
				str(coord1[1]+1) + ":"+sheet2 + "." + \
                                _digit_to_alpha(coord1[0]) + str(coord2[1]+1)

	    my_dict["data"] = sheet1 + "."+ _digit_to_alpha(coord1[0]+1) + \
			      str(coord1[1]+1) + ":"+sheet2 + "." + \
                              _digit_to_alpha(coord2[0]) + str(coord2[1]+1)


    return my_dict
    

def divide_range(cell_range, by="columns"):
    """
    Returns a list of cell range (string) divided in columns or rows

    cell_range -- str like Sheet.A1:Sheet.C5
    by -- 'columns' 'rows'
    """
    my_list = []
    sheet1=splitfields(splitfields(cell_range, ":")[0],".")[0]
    cell1=splitfields(splitfields(cell_range, ":")[0],".")[1]
    sheet2=splitfields(splitfields(cell_range, ":")[1],".")[0]
    cell2=splitfields(splitfields(cell_range, ":")[1],".")[1]
    
    tmp_coord1 = _get_cell_coordinates(cell1)
    tmp_coord2 = _get_cell_coordinates(cell2)

    #we need the coordinate in a crescent order
    coord1 = (min(tmp_coord1[0], tmp_coord2[0]),min(tmp_coord1[1],
                                                                tmp_coord2[1]))

    coord2 = (max(tmp_coord1[0], tmp_coord2[0]),max(tmp_coord1[1],
                                                                tmp_coord2[1]))


    if by == 'columns':
        for i in range(coord1[0], coord2[0]+1):
	    tmp_cell_range = sheet1 + "." + _digit_to_alpha(i) + \
			     str(coord1[1]+1) + ":" + sheet2 + "." + \
                             _digit_to_alpha(i) + str(coord2[1]+1)

            my_list.append(tmp_cell_range)
        return my_list
        
    elif by == 'rows':
        for i in range(coord1[1], coord2[1]+1):
	    tmp_cell_range = sheet1 + "." + _digit_to_alpha(coord1[0]) + \
			     str(i+1) + ":" + sheet2 + "." + \
                             _digit_to_alpha(coord2[0]) + str(i+1) \

            my_list.append(tmp_cell_range)
        return my_list

    else:
        raise AttributeError

def add_chart_structure_in_document(document):
    """Search the .otc template installed and put xml files in document """
    #We search the template
    my_template = ""
    for p in path:
        if exists(p+"/chart/templates/chart.otc"):
            my_template=p+"/chart/templates/chart.otc"
            break
    
    if my_template == "":
        raise IOError, "Template .otc not found"
    
    #we have to create a folder for the chart
    i = 1
    obj_created = False
    name=""
    manifest = document.get_part('META-INF/manifest.xml')
    while not obj_created:
        name="Object "+str(i)
        if manifest.get_media_type(name+'/') is None:
	    manifest.add_full_path(name+'/',
                                    "application/vnd.oasis.opendocument.chart")

            document.container.set_part(name+'/', '')
            obj_created = True
        else:
            i = i+1

    #We open the template
    chart = odf_new_container(my_template)

    #we add templates files in the document
    document.set_part(name+'/content.xml', chart.get_part('content.xml'))
    manifest.add_full_path(name+'/content.xml' , "text/xml")

    document.set_part(name+'/styles.xml', chart.get_part('styles.xml'))
    manifest.add_full_path(name+'/styles.xml' , "text/xml")
    
    document.set_part(name+'/meta.xml', chart.get_part('meta.xml'))
    manifest.add_full_path(name+'/meta.xml' , "text/xml")

    return name

def attach_chart_to_cell(name_obj, cell):
    """
    create a frame in 'cell' in order to display the chart 'name_obj'
    name_obj - str
    cell - odf_cell
    """
    #We need a frame
    frame = odf_create_frame(size=("10cm", "10cm"))

    #We need a draw:object element
    element = odf_create_element("draw:object")
    element.set_attribute("xlink:href", "./"+name_obj)
    element.set_attribute("xlink:type", "simple")
    element.set_attribute("xlink:show", "embed")
    element.set_attribute("xlink:actuate", "onLoad")
                     
    frame.append(element)
    cell.append(frame)

    return cell



#register
register_element_class('chart:chart', odf_chart)
register_element_class('chart:plot-area', odf_plot_area)
