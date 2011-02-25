###
#Import
###
import csv
import sys

#import lpod
from lpod.document import odf_new_document
from lpod.table import odf_create_table, odf_create_row, odf_create_cell
from lpod.table import _digit_to_alpha
from lpod.xmlpart import odf_xmlpart

#import lpod_vle
from lpod_chart.chart_style import odf_create_chart_title_style 
from lpod_chart.chart_style import odf_create_chart_legend_style
from lpod_chart.chart_style import odf_create_chart_plot_area_style
from lpod_chart.chart_style import odf_create_chart_axis_style
from lpod_chart.chart_style import odf_create_chart_axis_title_style
from lpod_chart.chart_style import odf_create_chart_wall_style
from lpod_chart.chart_style import odf_create_chart_style
from lpod_chart.chart import create_simple_chart, attach_chart_to_cell
from lpod_chart.chart import divide_range
from lpod_chart.chart import add_chart_structure_in_document

palette=['#0000ff', '#bf00ff', '#ff0080', '#ff4000', '#ffff00', '#40ff00',
'#00ff7f', '#00bfff', '#6000ff', '#ff00df', '#ff0020', '#ff9f00', '#9fff00',
'#00ff20', '#00ffdf', '#0060ff']

###
#First step : we read the data and convert it the way it will give us the chart we want
###
if (len(sys.argv) > 2):

    m_file = open(sys.argv[1])
    m_reader = csv.reader(m_file, delimiter='\t')

    #we read and keep all the data from m_file
    data = []
    for row in m_reader:
        data.append(row)


    #we search how many rows (representing a place) we'll need
    places = []
    places_book = {}
    address = 0
    places_status = {}
    for i in range(len(data)):
        if data[i][0] != 'DD':
            if data[i][2] not in places:
                places.append(data[i][2])
                places_book[data[i][2]] = address
                address = address +1
                places_status[data[i][2]]=0
                

    line_read = []
    result= []
    colors = []
    while len(line_read) < len(data)-1:
        #we need the minimum for the first value
        minimum = 1000000.0
        index = 0
        for i in range(len(data)):
	    if data[i][0] != 'DD'and float(data[i][0]) < minimum and i not in \
                                                                   line_read:

                minimum = float(data[i][0])
                index = i

        #we add the two data in result : the waiting and the activity
        beginning = float(data[index][0])
        end = float(data[index][1])
        place_involved = int(data[index][2])
        colors.append(data[index][3])

        tmp1=[]
        for i in places:
            tmp1.append(0)

        tmp2=[]
        for i in places:
            tmp2.append(0)

	tmp1[places_book[data[index][2]]] = beginning - \
                                                 places_status[data[index][2]]

        result.append(tmp1)
        tmp2[places_book[data[index][2]]] = end - beginning
        result.append(tmp2)

        #we update the status
	places_status[data[index][2]] = places_status[data[index][2]] + \
                   (beginning - places_status[data[index][2]])+(end - beginning)

        line_read.append(index)

###
#Second step :  write the data in an ods file
###

    document = odf_new_document('spreadsheet')
    body = document.get_body()      

    table = odf_create_table('Data')
    body.append(table)

    for i in range(len(result[0])):
        row = odf_create_row()
        cell = odf_create_cell("Place "+str(places[i]))
        row.append_cell(cell)
        for j in range(len(result)):
            cell = odf_create_cell(result[j][i])
            row.append_cell(cell)
        table.append_row(row)

###
#Third step :  create our chart
###
    table = odf_create_table('Gantt')
    body.append(table)
    
    cell_range="Data.A1:Data."+_digit_to_alpha(len(result))+str(len(result[0]))

    cols_list = divide_range(cell_range)

    #we add the xml files that will contain the chart informations
    chart_directory = add_chart_structure_in_document(document)

    #we need an access to the office:automatic-styles element
    chart_content = odf_xmlpart(chart_directory+'/content.xml', document)
    styles =  chart_content.get_element("office:automatic-styles")

    #we add informations about style
    styles.append(odf_create_chart_title_style("titre", "12pt"))
    plot_s = odf_create_chart_plot_area_style("plot")
    plot_s.set_vertical("true")
    plot_s.set_stacked("true")
    styles.append(plot_s)

    axe_x_style = odf_create_chart_axis_style("axe x")
    styles.append(axe_x_style)

    axe_y_style = odf_create_chart_axis_style("axe y")
    axe_y_style.set_properties({'chart:minimum': '0'}, area='chart')
    styles.append(axe_y_style)

    styles.append(odf_create_chart_axis_style("axe y"))
    styles.append(odf_create_chart_wall_style("wall"))

    #invisible style for the the waiting data
    s = odf_create_chart_style("invisible")
    s.set_fill("none", "#FFFFFF")
    s.set_stroke("none", "#FFFFFF")
    styles.append(s)


    #others styles for activity
    i=0
    for c in palette:
        s = odf_create_chart_style("colored"+str(i))
        s.set_fill("solid", c)
        styles.append(s)
        i = i + 1

    #we need an access to the office:chart element
    chart =  chart_content.get_element("office:body/office:chart")

    #we create our chart
    c = create_simple_chart('bar', cell_range, 'Gantt', data_in_columns=True)
    c.delete(c.get_element("chart:legend"))
    a = c.get_element("chart:plot-area")
    a.set_categories('x', "Data.A1:Data.A"+str(len(result[0])))

    #then we attach style informations with the element associated
    c.set_title_style("titre")
    a.set_style("plot")
    a.set_axis_style('x', 'axe x')
    a.set_axis_style('y', 'axe y')
    a.set_wall(style="wall")

    i=1 
    for col in cols_list:
        if i % 2 == 0:
            a.set_chart_series_style(col, "invisible")
        else:
            num = int(colors[i/2-1])%16
            a.set_chart_series_style(col, "colored"+str(num))
        i=i+1

    #we put our chart (named 'c') in the office:chart element
    chart.append(c)
    #we need to rewrite the content.xml with all the new informations
    document.set_part(chart_directory+'/content.xml', chart_content.serialize())

    #we need to attach the chart with a cell
    row = odf_create_row()
    cell = odf_create_cell()

    #'cell' will contain the frame with our chart 
    cell = attach_chart_to_cell(chart_directory, cell)

    row.append_cell(cell)
    table.append_row(row)
    body.append(table)

###
#Finally we hide the first sheet
###

    content = document.get_part('content.xml')
    style = content.get_style("table", "ta1")
    p = style.set_properties({'table:display': 'false'})
    t = content.get_element("//table:table[@table:name='Data']")
    t.set_attribute("table:style-name", "ta1")


    document.save(sys.argv[2])

else:
    print "Two arguments needed :"
    print "You have to give file with the data for gantt and the ods output"
    print "example : gantt_example.py gantt_example_data out.ods"
        
