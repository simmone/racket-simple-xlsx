#lang scribble/manual

@title{Set Styles}

Add styles to sheets.

@section{about styles}

1. styles only can be setted in one sheet, need be in a  with-data-sheet-*'s scope.

2. you can set cell range, row range, col range styles.

3. if you have overlap styles, the overlap area's style will be piled up.

@section{set-row-range-height}

@codeblock|{set-row-range-height (-> string? natural? void?)}|

arg1: row range, as "1-2".

arg2: row height.

@section{set-col-range-width}

@codeblock|{set-col-range-width (-> string? natural? void?)}|

arg1: col range, as "A-B".

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

@section{set-cell-range-border-style}

@codeblock|{set-cell-range-border-style (-> string? border-direction? rgb? border-mode? void?)}|

arg1: cell range, as "A1-B2".

arg2: border-direction?, one of '("all" "side" "top" "bottom" "left" "right").
the side direction means only set the cell range's out border.

arg3: rgb?, rgb color as "0000ff".

arg4: border-mode?, one of '("" "thin" "dashed" "double" "thick")

@codeblock|{
(set-cell-range-border-style "B2-F6" "all" "ff0000" "thick")
(set-cell-range-border-style "B8-F12" "left" "ff0000" "thick")
(set-cell-range-border-style "H2-L6" "right" "ff0000" "dashed")
(set-cell-range-border-style "H8-L12" "top" "ff0000" "double")
(set-cell-range-border-style "N2-R6" "bottom" "ff0000" "thick")
(set-cell-range-border-style "N8-R12" "side" "ff0000" "thick")
}|

@image{scribble/style/border_style.png}

@section{set-cell-range-font-style}

@codeblock|{set-cell-range-font-style (-> string? natural? string? rgb? void?)}|

arg1: cell range.

arg2: font size.

arg3: font name, as "Arial".

arg4: font color, rgb color, as "0000ff".

@codeblock|{
(set-cell-range-font-style "A1-C1" 12 "Arial" "000000")
(set-cell-range-font-style "A2-C2" 16 "Monospace" "900000")
(set-cell-range-font-style "A3-C3" 20 "Sans" "990000")
}|

@image{scribble/style/font_style.png}

@section{set-row-range-font-style}

@codeblock|{set-row-range-font-style (-> string? natural? string? rgb? void?)}|

arg1: row range.

@section{set-col-range-font-style}

@codeblock|{set-col-range-font-style (-> string? natural? string? rgb? void?)}|

arg1: col range.

@section{set-cell-range-alignment-style}

@codeblock|{set-cell-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)}|

arg1: cell range.

arg2: horizontal_mode?, one of '("left" "right" "center")

arg3: vertical_mode?, one of '("top" "bottom" "center")

@codeblock|{
(set-cell-range-alignment-style "A1-E5" "center" "center")
(set-cell-range-border-style "A1-E5" "side" "ff0000" "thick")

(set-cell-range-alignment-style "G1-K5" "left" "top")
(set-cell-range-border-style "G1-K5" "side" "ff0000" "thick")

(set-cell-range-alignment-style "M1-Q5" "right" "bottom")
(set-cell-range-border-style "M1-Q5" "side" "ff0000" "thick")

(set-row-range-height "1-5" 30)
}|

@image{scribble/style/alignment_style.png}

@section{set-row-range-alignment-style}

arg1: row range.

@codeblock|{set-row-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)}|

@section{set-col-range-alignment-style}

arg1: col range.

@codeblock|{set-col-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)}|

@section{set-cell-range-number-style}

@codeblock|{set-cell-range-number-style (-> string? string? void?)}|

arg1: cell range.

arg2: number style as "0.00" "0,000.00" "0.00%" etc.

@codeblock|{
(set-cell-range-number-style "A1-C1" "0.00")
(set-cell-range-number-style "A2-C2" "0.000")
(set-cell-range-number-style "A3-C3" "0,000.00%")
(set-col-range-width "A-C" 30)
(set-row-range-height "1-3" 50)
}|

@image{scribble/style/number_style.png}

@section{set-row-range-number-style}

@codeblock|{set-row-range-number-style (-> string? string? void?)}|

arg1: row range.

@section{set-col-range-number-style}

@codeblock|{set-col-range-number-style (-> string? string? void?)}|

arg1: col range.

@section{set-cell-range-date-style}

@codeblock|{set-cell-range-date-style (-> string? string? void?)}|

arg1: cell range.

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

@section{set-row-range-date-style}

arg1: row range.

@section{set-col-range-date-style}

arg1: col range.

          [set-row-range-date-style (-> string? string? void?)]
          [set-col-range-date-style (-> string? string? void?)]

@section{set-cell-range-fill-style}

@codeblock|{set-cell-range-fill-style (-> string? rgb? fill-pattern? void?)}|

arg1: cell range.

arg2: rgb color as "0000ff".

arg3: fill pattern, one of 
@codeblock|{
'("solid" "gray125" "darkGray" "mediumGray" "lightGray"
"gray0625" "darkHorizontal" "darkVertical" "darkDown" "darkUp"
"darkGrid" "darkTrellis" "lightHorizontal" "lightVertical" "lightDown"
"lightUp" "lightGrid" "lightTrellis")
}|

@codeblock|{
(set-cell-range-fill-style "B2-F6" "ff0000" "solid")
(set-cell-range-fill-style "H2-L6" "0000ff" "gray125")
(set-cell-range-fill-style "N2-R6" "00ff00" "darkDown")
}|

@image{scribble/style/fill.png}

@section{set-row-range-fill-style}

@codeblock|{set-row-range-fill-style (-> string? rgb? fill-pattern? void?)}|

arg1: row range.

@section{set-col-range-fill-style}

@codeblock|{set-col-range-fill-style (-> string? rgb? fill-pattern? void?)}|

arg1: col range.