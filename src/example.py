from lpod.document import odf_new_document
from lpod.table import odf_create_table
from random import randrange

#We create a spreadsheet just for the example
doc = odf_new_document('spreadsheet')
body = doc.get_body()
table = odf_create_table("Data", width=5, height=5)
body.append(table)
c=0
l=0
for c in range(5):
    for l in range(5):
        v = randrange(0,1000)
        table.set_value((c, l), v)


#we add the xml files that will contain the chart informations
from chart import add_chart_structure_in_document
chart_directory = add_chart_structure_in_document(doc)

#we need an access to the office:automatic-styles element
from lpod.xmlpart import odf_xmlpart
chart_content = odf_xmlpart(chart_directory+'/content.xml', doc)
styles =  chart_content.get_element("office:automatic-styles")

#we add informations about style
from chart_style import odf_create_chart_title_style
from chart_style import odf_create_chart_legend_style
from chart_style import odf_create_chart_plot_area_style
from chart_style import odf_create_chart_axis_style
from chart_style import odf_create_chart_axis_title_style
from chart_style import odf_create_chart_wall_style, odf_create_chart_style

styles.append(odf_create_chart_title_style("titre", "12pt"))
plot_s=odf_create_chart_plot_area_style("plot")
styles.append(plot_s)
styles.append(odf_create_chart_axis_style("axe x"))
styles.append(odf_create_chart_axis_style("axe y"))
styles.append(odf_create_chart_wall_style("wall"))

#we choose a color for each series
palette=['#0000ff', '#bf00ff', '#ff0080', '#ff4000', '#ffff00', '#40ff00',
'#00ff7f', '#00bfff', '#6000ff', '#ff00df', '#ff0020', '#ff9f00', '#9fff00',
'#00ff20', '#00ffdf', '#0060ff']

from chart import divide_range
cols_list = divide_range("Data.A1:Data.E5")
for i in range(len(cols_list)):
    s = odf_create_chart_style("series"+str(i))
    s.set_fill("solid", palette[i])
    styles.append(s)

#we need an access to the office:chart element
chart =  chart_content.get_element("office:body/office:chart")

#we create our chart
from chart import create_simple_chart
c = create_simple_chart('line', 'Data.A1:Data.E5', 'Mon titre',
                                                         data_in_columns=True)

a = c.get_element("chart:plot-area")

#then we attach style informations with the element associated
c.set_title_style("titre")
a.set_style("plot")
a.set_axis_style('x', 'axe x')
a.set_axis_style('y', 'axe y')
a.set_wall(style="wall")
i=0
for col in cols_list:
    a.set_chart_series_style(col, "series"+str(i))
    i=i+1

#we put our chart (named 'c') in the office:chart element
chart.append(c)
#we need to rewrite the content.xml with all the new informations
doc.set_part(chart_directory+'/content.xml', chart_content.serialize())

#we need to attach the chart with a cell
from lpod.table import odf_create_row, odf_create_cell
row = odf_create_row()
cell = odf_create_cell()

#'cell' will contain the frame with our chart 
from chart import attach_chart_to_cell
cell = attach_chart_to_cell(chart_directory, cell)

row.append_cell(cell)
table.append_row(row)
body.append(table)

#we save our document
doc.save("out.ods")
