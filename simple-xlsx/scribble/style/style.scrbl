#lang scribble/manual

@title{Set Styles}

add styles to sheets.

@section{about styles}

1. styles only can be setted in one sheet, need be in a  with-data-sheet-*'s scope.

2. you can set cell range, row range, col range styles.

3. if you have overlap styles, the overlap area's style will be piled up.

@section{cell/row/col range}

cell range: "A1-B3".

row range: "1-3", start from 1.

col range: "A-C" or "1-3", start from 1.

@section{set-row-range-height}

@codeblock|{set-row-range-height (-> string? natural? void?)}|

arg1: row range.

arg2: row height.

@section{set-col-range-width}

@codeblock|{set-col-range-width (-> string? natural? void?)}|

arg1: col range.

arg2: col width.

Example:

@codeblock|{
(set-col-range-width "A-C" 30)
(set-row-range-height "1-2" 40)
}|

@image{scribble/style/row_width_and_col_height.png}

@section{set-freeze-row-col-range}

@codeblock|{set-freeze-row-col-range (-> natural? natural? void?)}|

freeze rows and cols.

arg1: rows count.

arg2: cols count.

@codeblock|{
(set-freeze-row-col-range 2 2)
}|

@section{set-merge-cell-range}

@codeblock|{set-merge-cell-range (-> cell-range? void?)}|

@codeblock|{cell-range?: as "A1-C3" or "A1:C3"}|

set merge cell ranges.(multiple times)

@codeblock|{
(set-merge-cell-range "A1-C3")
(set-merge-cell-range "D5-F7")
(set-merge-cell-range "G8-I10")
}|

@section{about color}

you can set rgb color, like: FF0000.

1. RGB use upcase, not support lower case string.

2. Not support theme color yet, only support standard color.
   If sheet software set theme color, write back will lost color information.

@section{set-cell-range-border-style}

@codeblock|{
  set-cell-range-border-style (-> string? border-direction? rgb? border-mode? void?)
  set-row-range-border-style (-> string? border-direction? rgb? border-mode? void?)
  set-col-range-border-style (-> string? border-direction? rgb? border-mode? void?)
}|

arg1: cell/row/col range.

arg2: border-direction?, one of '("all" "side" "top" "bottom" "left" "right").
the side direction means only set the cell range's out border.

arg3: rgb?, rgb color as "0000FF".

arg4: border-mode?, one of '("" "thin" "dashed" "double" "thick")

@codeblock|{
(set-cell-range-border-style "B2-F6" "all" "FF0000" "thick")
(set-cell-range-border-style "B8-F12" "left" "FF0000" "thick")
(set-cell-range-border-style "H2-L6" "right" "FF0000" "dashed")
(set-cell-range-border-style "H8-L12" "top" "FF0000" "double")
(set-cell-range-border-style "N2-R6" "bottom" "FF0000" "thick")
(set-cell-range-border-style "N8-R12" "side" "FF0000" "thick")
}|

@image{scribble/style/border_style.png}

@section{set-font-style}

@codeblock|{
  set-cell-range-font-style (-> string? natural? string? rgb? void?)
  set-row-range-font-style (-> string? natural? string? rgb? void?)
  set-col-range-font-style (-> string? natural? string? rgb? void?)
}|

arg1: cell/row/col range.

arg2: font size.

arg3: font name, as "Arial".

arg4: font color, rgb color, as "0000FF".

@codeblock|{
(set-cell-range-font-style "A1-C1" 12 "Arial" "000000")
(set-cell-range-font-style "A2-C2" 16 "Monospace" "900000")
(set-cell-range-font-style "A3-C3" 20 "Sans" "990000")
}|

@image{scribble/style/font_style.png}

@section{set-alignment-style}

@codeblock|{
  set-cell-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)
  set-row-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)
  set-col-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)
}|

arg1: cell/row/col range.

arg2: horizontal_mode?, one of '("left" "right" "center")

arg3: vertical_mode?, one of '("top" "bottom" "center")

@codeblock|{
(set-cell-range-alignment-style "A1-E5" "center" "center")
(set-cell-range-border-style "A1-E5" "side" "FF0000" "thick")

(set-cell-range-alignment-style "G1-K5" "left" "top")
(set-cell-range-border-style "G1-K5" "side" "FF0000" "thick")

(set-cell-range-alignment-style "M1-Q5" "right" "bottom")
(set-cell-range-border-style "M1-Q5" "side" "FF0000" "thick")

(set-row-range-height "1-5" 30)
}|

@image{scribble/style/alignment_style.png}

@section{set-number-style}

@codeblock|{
  set-cell-range-number-style (-> string? string? void?)
  set-row-range-number-style (-> string? string? void?)
  set-col-range-number-style (-> string? string? void?)
}|

arg1: cell/row/col range.

arg2: number style as "0.00" "0,000.00" "0.00%" etc.

@codeblock|{
(set-cell-range-number-style "A1-C1" "0.00")
(set-cell-range-number-style "A2-C2" "0.000")
(set-cell-range-number-style "A3-C3" "0,000.00%")
(set-col-range-width "A-C" 30)
(set-row-range-height "1-3" 50)
}|

@image{scribble/style/number_style.png}

@section{set-date-style}

@codeblock|{
  set-cell-range-date-style (-> string? string? void?)
  set-row-range-date-style (-> string? string? void?)
  set-col-range-date-style (-> string? string? void?)
}|

arg1: cell/row/col range.

arg2: date style, as "yyyy/mm/dd", "yyyy-mm-dd", "yyyymmdd" etc.

@codeblock|{
...
(add-data-sheet
  "Sheet1"
  (list
    (list 
      (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
      (seconds->date (find-seconds 0 0 0 17 9 2018 #f))
      (seconds->date (find-seconds 0 0 0 17 9 2018 #f)))))

...

(set-cell-range-date-style "A1" "yyyy-mm-dd")
(set-cell-range-date-style "B1" "yyyy/mm/dd")
(set-cell-range-date-style "C1" "yyyymmdd")

(set-col-range-width "A-C" 20)
(set-row-range-height "1-3" 20)
}|

@image{scribble/style/date_style.png}

@section{set-fill-style}

@codeblock|{
  set-cell-range-fill-style (-> string? rgb? fill-pattern? void?)
  set-row-range-fill-style (-> string? rgb? fill-pattern? void?)
  set-col-range-fill-style (-> string? rgb? fill-pattern? void?)
}|

arg1: cell/row/col range.

arg2: rgb color as "0000FF".

arg3: fill pattern, one of 
@codeblock|{
'("solid" "gray125" "darkGray" "mediumGray" "lightGray"
"gray0625" "darkHorizontal" "darkVertical" "darkDown" "darkUp"
"darkGrid" "darkTrellis" "lightHorizontal" "lightVertical" "lightDown"
"lightUp" "lightGrid" "lightTrellis")
}|

@codeblock|{
(set-cell-range-fill-style "B2-F6" "FF0000" "solid")
(set-cell-range-fill-style "H2-L6" "0000FF" "gray125")
(set-cell-range-fill-style "N2-R6" "00FF00" "darkDown")
}|

@image{scribble/style/fill.png}

