#lang scribble/manual

@title{Add Chart Sheet}

add chart sheet to xlsx.

@section{add-chart-sheet}

@codeblock|{
[add-chart-sheet (-> string?
                 (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D)
                 string?
                 (listof (list/c string? string? string? string? string?)) void?)]
}|

arg1: chart sheet name.

arg2: chart type, one of ('LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D).

arg3: chart topic, will display on the above top.

arg4: chart data serial, serial list.

serial arguments:

arg1 string: category name.

arg2 string: category data sheet name.

arg3 string: category data range.

arg4 string: value data sheet name.

arg5 string: value data range.

@section{Line Chart}

@codeblock|{
(add-data-sheet 
  "DataSheet"
  '(
    ("201601" "201602" "201603" "201604")
    (100 300 200 400)
    (200 400 300 100)
    (300 500 400 200)
))

(add-chart-sheet
  "LineChart" 'LINE "LineChartExample"
  '(
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
    ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")
    ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
))
}|

@image{scribble/chartsheet/line_chart.png}

@codeblock|{
(add-chart-sheet
  "Line3DChart" 'LINE3D "Line3DChartExample"
  '(
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
    ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")
    ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
))
}|

@image{scribble/chartsheet/line_3d_chart.png}

@section{Bar Chart}

@codeblock|{
(add-chart-sheet
  "BarChart" 'BAR "BarChartExample"
  '(
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
    ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")
    ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
))
}|

@image{scribble/chartsheet/bar_chart.png}

@codeblock|{
(add-chart-sheet
  "Bar3DChart" 'BAR3D "Bar3DChartExample"
  '(
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
    ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")
    ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
))
}|

@image{scribble/chartsheet/bar_3d_chart.png}

@section{Pie Chart}

@codeblock|{
(add-chart-sheet
  "PieChart" 'PIE "PieChartExample"
  '(
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
))
}|

@image{scribble/chartsheet/pie_chart.png}

@codeblock|{
(add-chart-sheet
  "Pie3DChart" 'PIE3D "Pie3DChartExample"
  '(
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
))
}|

@image{scribble/chartsheet/pie_3d_chart.png}