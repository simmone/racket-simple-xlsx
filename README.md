# simple-xlsx: access .xlsx file in Racket

Chen Xiao <[chenxiao770117@gmail.com](mailto:chenxiao770117@gmail.com)>

The `simple-xlsx` read and write `.xlsx` file.

1. Campatible with Microsoft Office, Google Sheets, Libre Office, WPS.

2. Support string, number, date as data type.

3. Multiple style supported: color, border, font, etc.

4. Can generate simple chart sheet.

    1 Install                                              
                                                           
    2 Basic Example                                        
      2.1 Generate a xlsx file                             
      2.2 Read a xlsx file                                 
      2.3 Read a xlsx file, modify something and Write back
                                                           
    3 Top Level: \*-xlsx                                   
      3.1 write-xlsx                                       
      3.2 read-xlsx                                        
      3.3 read-and-write-xlsx                              
      3.4 get-sheet-name-list                              
      3.5 get-sheet-count                                  
                                                           
    4 Second Level: with-sheet\*                           
      4.1 with-sheet-\*                                    
                                                           
    5 Add Data Sheet                                       
      5.1 add-data-sheet                                   
                                                           
    6 Access Data                                          
      6.1 cell/cell-value?                                 
      6.2 Example Sheet Data                               
      6.3 Skip the with-sheet\*                            
      6.4 get-cell/set-cell!                               
      6.5 get-row/set-row!                                 
      6.6 get-rows/set-rows!                               
      6.7 get-col/set-col!                                 
      6.8 get-cols/set-cols!                               
      6.9 get-rows-count/get-cols-count                    
                                                           
    7 Set Styles                                           
      7.1 about styles                                     
      7.2 cell/row/col range                               
      7.3 set-row-range-height                             
      7.4 set-col-range-width                              
      7.5 set-freeze-row-col-range                         
      7.6 set-merge-cell-range                             
      7.7 about color                                      
      7.8 set-cell-range-border-style                      
      7.9 set-font-style                                   
      7.10 set-alignment-style                             
      7.11 set-number-style                                
      7.12 set-date-style                                  
      7.13 set-fill-style                                  
                                                           
    8 Add Chart Sheet                                      
      8.1 add-chart-sheet                                  
      8.2 Line Chart                                       
      8.3 Bar Chart                                        
      8.4 Pie Chart                                        

## 1. Install

raco pkg install simple-xlsx

Caution: simple-xlsx depends on package: simple-xml, if you have
installed simple-xml, should update the package to newest version.

## 2. Basic Example

### 2.1. Generate a xlsx file

`(write-xlsx`                                                       
  `"basic.xlsx"`                                                    
  `(lambda ()`                                                      
    `(add-data-sheet "Sheet1" '(("month1" "month2" "month3" "month4"
"real")))`                                                          
                                                                    
    `(add-data-sheet "Sheet2" '((201601 100 110 1110 6.9)))))`      

1. All operations in `write-xlsx`’s lambda scope.

2. Specify file name, sheet name, have same count’s list data, done.

### 2.2. Read a xlsx file

`(read-xlsx`                                                          
  `"basic_write.xlsx"`                                                
  `(lambda ()`                                                        
    `(check-equal? (get-sheet-name-list) '("Sheet1" "Sheet2"))`       
                                                                      
    `(with-sheet-ref`                                                 
    `0`                                                               
    `(lambda ()`                                                      
      `(check-equal? (get-row 1) '("month1" "month2" "month3" "month4"
"real"))))`                                                           
                                                                      
    `(with-sheet-ref`                                                 
    `1`                                                               
    `(lambda ()`                                                      
      `(check-equal? (get-row 1) '(201601 100 110 1110 6.9))))))`     

Navigate to a specific sheet have two ways: use index or name

1. use sheet index: \*\*\*\*-ref, index from 0

`(with-sheet-ref`   
  `sheet_index`     
  `(lambda () ...))`

2. use sheet name: \*\*\*\*-name

`(with-sheet-name`  
  `sheet_name`      
  `(lambda () ...))`

### 2.3. Read a xlsx file, modify something and Write back

`(read-and-write-xlsx`                                              
  `basic_write_file`                                                
  `basic_read_and_write_file`                                       
  `(lambda ()`                                                      
    `(check-equal? (get-sheet-name-list) '("Sheet1" "Sheet2"))`     
                                                                    
    `(with-sheet-ref`                                               
    `0`                                                             
    `(lambda ()`                                                    
      `(set-cell-value! "B1" "John")`                               
      `(check-equal? (get-row 1) '("month1" "John" "month3" "month4"
"real"))))`                                                         
                                                                    
    `(with-sheet-ref`                                               
    `1`                                                             
    `(lambda ()`                                                    
      `(check-equal? (get-row 1) '(201601 100 110 1110 6.9))))`     
      `))`                                                          

The first arg is read file, second is the write back file, these two can
be a same file, if you want to replace the oringinal.

## 3. Top Level: \*-xlsx

All the operations on the xlsx file: write, read, modify, should be
placed in top level functions:

write-xlsx: In its scope, add sheets, set data or styles, in the end,
generate a new xlsx file.

read-xlsx: In its scope, read from a xlsx file, get data.

read-and-write-xlsx: In its scope, read from a xlsx file, set data or
styles, in the end, write back a new file or overlap the original file.

### 3.1. write-xlsx

```racket
write-xlsx (-> path-string? procedure? any)
```

arg1: The output file name.

arg2: user procedure.

### 3.2. read-xlsx

```racket
read-xlsx (-> path-string? procedure? any)
```

arg1: The input file name.

arg2: user procedure.

### 3.3. read-and-write-xlsx

```racket
read-and-write-xlsx (-> path-string? path-string? procedure? any)
```

arg1: The input file name.

arg2: The output file name.

arg3: user procedure.

### 3.4. get-sheet-name-list

```racket
get-sheet-name-list (-> (listof string?))
```

### 3.5. get-sheet-count

```racket
get-sheet-count (-> natural?)
```

## 4. Second Level: with-sheet\*

In the with-sheet-\*’s scope, you can get and set sheet data, set sheet
styles, etc.

### 4.1. with-sheet-\*

```racket
with-sheet-ref (-> natural? procedure? any)
with-sheet (-> procedure? any)             
with-sheet-name (-> string? procedure? any)
```

with-sheet-ref’s first argument is the sheet index, start from 0.
with-sheet means with-sheet-ref 0. with-sheet-name use sheet name to
specify.

All the sheet data’s operations: get or set data, set data’s styles
should be placed in the with-sheet-\*.

Because most methods is effect in the sheet scope, so normally, the code
style is like below:

```racket
(write-xlsx                                                                  
  "out.xlsx"                                                                 
  (lambda ()                                                                 
    (add-data-sheet "Sheet1" '(("month1" "month2" "month3" "month4" "real")))
                                                                             
    (with-sheet-ref                                                          
    0                                                                        
    (lambda ()                                                               
      (set-cell! "B1" "John")))))                                            
```

## 5. Add Data Sheet

In the write-xlsx, you can add data sheets.

### 5.1. add-data-sheet

```racket
add-data-sheet (-> string? (listof list?) void?)
```

first argument is sheet name: sheet name is unique, can’t have
duplicated sheet names.

second argument is a listof list.

data type you can use: string, number, date.

data is a listof list, each list’length can be different.

according to the longest list, function will pad value on the right to
keep all the list have the same length.

default pad value is "", you can use \#:fill? to specify other values.

ie:

```racket
(add-data-sheet               
  "Sheet1"                    
  '(                          
    ("a" "b" "c")             
    (1)                       
    (1.0 2.0)                 
    ))                        
                              
 will add the data list below:
 '(                           
   ("a" "b" "c")              
   (1 "" "")                  
   (1.0 2.0 "")               
   )                          
```

use \#:start\_cell? to specify the datalist’s start cell, default is
"A1".

combine write-xlsx and add-data-sheet, you can generate a xlsx file:

```racket
(write-xlsx                                         
  basic_file                                        
  (lambda ()                                        
    (add-data-sheet                                 
      "Sheet1"                                      
      '(                                            
        ("month1" "month2" "month3" "month4" "real")
        (201601 100 110 1110 6.9))                  
       )                                            
     ))                                             
```

## 6. Access Data

### 6.1. cell/cell-value?

cell is the XLSX’s cell’s name:

First is the column’s index identified by the alphabet from "A".

Second is the row’s index identified by the number from 1.

Example: Cell in the second row, the third column, so cell is "C2".

cell-value can be these type: string, number, date.

### 6.2. Example Sheet Data

```racket
(add-data-sheet "sheet1" '(        
                            (1 2 3)
                            (4 5 6)
                          ))       
```

### 6.3. Skip the with-sheet\*

If not have a mass access on a sheet, you can get/set data a sheet
directly, not need in a with-sheet\* scope.

Example:

1. Normally get a cell’s value:

```racket
(with-sheet (lambda () (get-cell "A1") ...))
```

2. Direct get a cell’s value:

```racket
(get-sheet-ref-cell 0 "A1")
```

There are 3 ways to locate a sheet:

sheet-ref: use sheet index, start from 0.

sheet-name: use sheet name, match exactly.

sheet-\*name\*: use part of sheet name, locate first matched sheet name.

Example:

get-cell have 3 direct function:

```racket
get-sheet-ref-cell, get-sheet-name-cell, get-sheet-*name*-cell
```

### 6.4. get-cell/set-cell!

```racket
get-cell (-> string? cell-value?)
                                 
(check-equals (get-cell "B2") 5) 
```

Direct function:

```racket
get-sheet-ref-cell, get-sheet-name-cell, get-sheet-*name*-cell
```

```racket
set-cell! (-> string? cell-value? void?)
                                        
(set-cell! "C1" 8)                      
```

Direct function:

```racket
set-sheet-ref-cell!, set-sheet-name-cell!, set-sheet-*name*-cell!
```

### 6.5. get-row/set-row!

Row’s index from 1.

```racket
get-row (-> natural? (listof cell-value?))
                                          
(check-equal? (get-row 1) '(1 2 3))       
```

Direct function:

```racket
get-sheet-ref-row, get-sheet-name-row, get-sheet-*name*-row
```

```racket
set-row! (-> natural? (listof cell-value?) void?)
                                                 
(set-row! 1 '(7 8 9))                            
```

Direct function:

```racket
set-sheet-ref-row!, set-sheet-name-row!, set-sheet-*name*-row!
```

### 6.6. get-rows/set-rows!

```racket
get-rows (-> (listof (listof cell-value?))) 
                                            
(check-equal? (get-rows) '((1 2 3) (4 5 6)))
```

Direct function:

```racket
get-sheet-ref-rows, get-sheet-name-rows, get-sheet-*name*-rows
```

```racket
set-rows! (-> (listof (listof cell-value?)) void?)
                                                  
(set-rows! '((1 2 3) (7 8 9)))                    
```

Direct function:

```racket
set-sheet-ref-rows!, set-sheet-name-rows!, set-sheet-*name*-rows!
```

### 6.7. get-col/set-col!

Col’s index from 1.

```racket
get-col (-> natural? (listof cell-value?))
                                          
(check-equal? (get-col 1) '(1 4))         
(check-equal? (get-col 2) '(2 5))         
(check-equal? (get-col 3) '(3 6))         
```

Direct function:

```racket
get-sheet-ref-col, get-sheet-name-col, get-sheet-*name*-col
```

```racket
set-col! (-> natural? (listof cell-value?) void?)
                                                 
(set-col! 1 '(7 8))                              
```

Direct function:

```racket
set-sheet-ref-col!, set-sheet-name-col!, set-sheet-*name*-col!
```

### 6.8. get-cols/set-cols!

```racket
get-cols (-> (listof (listof cell-value?)))   
                                              
(check-equal? (get-cols) '((1 4) (2 5) (3 6)))
```

Direct function:

```racket
get-sheet-ref-cols, get-sheet-name-cols, get-sheet-*name*-cols
```

```racket
set-cols! (-> (listof (listof cell-value?)) void?)
                                                  
(set-cols! '((7 8) (9 0) (1 2)))                  
```

Direct function:

```racket
set-sheet-ref-cols!, set-sheet-name-cols!, set-sheet-*name*-cols!
```

### 6.9. get-rows-count/get-cols-count

```racket
get-rows-count (-> natural?)     
get-cols-count (-> natural?)     
                                 
(check-equal? (get-rows-count) 2)
(check-equal? (get-cols-count) 3)
```

Direct function:

```racket
get-sheet-ref-rows-count, get-sheet-name-rows-count, get-sheet-*name*-rows-count
```

```racket
get-sheet-ref-cols-count, get-sheet-name-cols-count, get-sheet-*name*-cols-count
```

## 7. Set Styles

add styles to sheets.

### 7.1. about styles

1. styles only can be setted in one sheet, need be in a
with-data-sheet-\*’s scope.

2. you can set cell range, row range, col range styles.

3. if you have overlap styles, the overlap area’s style will be piled
up.

### 7.2. cell/row/col range

cell range: "A1-B3".

row range: "1-3", start from 1.

col range: "A-C" or "1-3", start from 1.

### 7.3. set-row-range-height

```racket
set-row-range-height (-> string? natural? void?)
```

arg1: row range.

arg2: row height.

### 7.4. set-col-range-width

```racket
set-col-range-width (-> string? natural? void?)
```

arg1: col range.

arg2: col width.

Example:

```racket
(set-col-range-width "A-C" 30) 
(set-row-range-height "1-2" 40)
```



### 7.5. set-freeze-row-col-range

```racket
set-freeze-row-col-range (-> natural? natural? void?)
```

freeze rows and cols.

arg1: rows count.

arg2: cols count.

```racket
(set-freeze-row-col-range 2 2)
```

### 7.6. set-merge-cell-range

```racket
set-merge-cell-range (-> cell-range? void?)
```

```racket
cell-range?: as "A1-C3" or "A1:C3"
```

set merge cell ranges.(multiple times)

```racket
(set-merge-cell-range "A1-C3") 
(set-merge-cell-range "D5-F7") 
(set-merge-cell-range "G8-I10")
```

### 7.7. about color

you can set rgb color, like: FF0000.

1. RGB use upcase, not support lower case string.

2. Not support theme color yet, only support standard color.    If sheet
software set theme color, write back will lost color information.

### 7.8. set-cell-range-border-style

```racket
set-cell-range-border-style (-> string? border-direction? rgb? border-mode? void?)
set-row-range-border-style (-> string? border-direction? rgb? border-mode? void?) 
set-col-range-border-style (-> string? border-direction? rgb? border-mode? void?) 
```

arg1: cell/row/col range.

arg2: border-direction?, one of ’\("all" "side" "top" "bottom" "left"
"right"\). the side direction means only set the cell range’s out
border.

arg3: rgb?, rgb color as "0000FF".

arg4: border-mode?, one of ’\("" "thin" "dashed" "double" "thick"\)

```racket
(set-cell-range-border-style "B2-F6" "all" "FF0000" "thick")   
(set-cell-range-border-style "B8-F12" "left" "FF0000" "thick") 
(set-cell-range-border-style "H2-L6" "right" "FF0000" "dashed")
(set-cell-range-border-style "H8-L12" "top" "FF0000" "double") 
(set-cell-range-border-style "N2-R6" "bottom" "FF0000" "thick")
(set-cell-range-border-style "N8-R12" "side" "FF0000" "thick") 
```



### 7.9. set-font-style

```racket
set-cell-range-font-style (-> string? natural? string? rgb? void?)
set-row-range-font-style (-> string? natural? string? rgb? void?) 
set-col-range-font-style (-> string? natural? string? rgb? void?) 
```

arg1: cell/row/col range.

arg2: font size.

arg3: font name, as "Arial".

arg4: font color, rgb color, as "0000FF".

```racket
(set-cell-range-font-style "A1-C1" 12 "Arial" "000000")    
(set-cell-range-font-style "A2-C2" 16 "Monospace" "900000")
(set-cell-range-font-style "A3-C3" 20 "Sans" "990000")     
```



### 7.10. set-alignment-style

```racket
set-cell-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?)
set-row-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?) 
set-col-range-alignment-style (-> string? horizontal_mode? vertical_mode? void?) 
```

arg1: cell/row/col range.

arg2: horizontal\_mode?, one of ’\("left" "right" "center"\)

arg3: vertical\_mode?, one of ’\("top" "bottom" "center"\)

```racket
(set-cell-range-alignment-style "A1-E5" "center" "center")   
(set-cell-range-border-style "A1-E5" "side" "FF0000" "thick")
                                                             
(set-cell-range-alignment-style "G1-K5" "left" "top")        
(set-cell-range-border-style "G1-K5" "side" "FF0000" "thick")
                                                             
(set-cell-range-alignment-style "M1-Q5" "right" "bottom")    
(set-cell-range-border-style "M1-Q5" "side" "FF0000" "thick")
                                                             
(set-row-range-height "1-5" 30)                              
```



### 7.11. set-number-style

```racket
set-cell-range-number-style (-> string? string? void?)
set-row-range-number-style (-> string? string? void?) 
set-col-range-number-style (-> string? string? void?) 
```

arg1: cell/row/col range.

arg2: number style as "0.00" "0,000.00" "0.00%" etc.

```racket
(set-cell-range-number-style "A1-C1" "0.00")     
(set-cell-range-number-style "A2-C2" "0.000")    
(set-cell-range-number-style "A3-C3" "0,000.00%")
(set-col-range-width "A-C" 30)                   
(set-row-range-height "1-3" 50)                  
```



### 7.12. set-date-style

```racket
set-cell-range-date-style (-> string? string? void?)
set-row-range-date-style (-> string? string? void?) 
set-col-range-date-style (-> string? string? void?) 
```

arg1: cell/row/col range.

arg2: date style, as "yyyy/mm/dd", "yyyy-mm-dd", "yyyymmdd" etc.

```racket
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
```



### 7.13. set-fill-style

```racket
set-cell-range-fill-style (-> string? rgb? fill-pattern? void?)
set-row-range-fill-style (-> string? rgb? fill-pattern? void?) 
set-col-range-fill-style (-> string? rgb? fill-pattern? void?) 
```

arg1: cell/row/col range.

arg2: rgb color as "0000FF".

arg3: fill pattern, one of

```racket
'("solid" "gray125" "darkGray" "mediumGray" "lightGray"               
"gray0625" "darkHorizontal" "darkVertical" "darkDown" "darkUp"        
"darkGrid" "darkTrellis" "lightHorizontal" "lightVertical" "lightDown"
"lightUp" "lightGrid" "lightTrellis")                                 
```

```racket
(set-cell-range-fill-style "B2-F6" "FF0000" "solid")   
(set-cell-range-fill-style "H2-L6" "0000FF" "gray125") 
(set-cell-range-fill-style "N2-R6" "00FF00" "darkDown")
```



## 8. Add Chart Sheet

add chart sheet to xlsx.

### 8.1. add-chart-sheet

```racket
[add-chart-sheet (-> string?                                                      
                 (or/c 'LINE 'LINE3D 'BAR 'BAR3D 'PIE 'PIE3D)                     
                 string?                                                          
                 (listof (list/c string? string? string? string? string?)) void?)]
```

arg1: chart sheet name.

arg2: chart type, one of (’LINE ’LINE3D ’BAR ’BAR3D ’PIE ’PIE3D).

arg3: chart topic, will display on the above top.

arg4: chart data serial, serial list.

serial arguments:

arg1 string: category name.

arg2 string: category data sheet name.

arg3 string: category data range.

arg4 string: value data sheet name.

arg5 string: value data range.

### 8.2. Line Chart

```racket
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
```



```racket
(add-chart-sheet                                      
  "Line3DChart" 'LINE3D "Line3DChartExample"          
  '(                                                  
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")   
    ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")  
    ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
))                                                    
```



### 8.3. Bar Chart

```racket
(add-chart-sheet                                      
  "BarChart" 'BAR "BarChartExample"                   
  '(                                                  
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")   
    ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")  
    ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
))                                                    
```



```racket
(add-chart-sheet                                      
  "Bar3DChart" 'BAR3D "Bar3DChartExample"             
  '(                                                  
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")   
    ("Puma" "DataSheet" "A1-D1" "DataSheet" "A3-D3")  
    ("Brooks" "DataSheet" "A1-D1" "DataSheet" "A4-D4")
))                                                    
```



### 8.4. Pie Chart

```racket
(add-chart-sheet                                   
  "PieChart" 'PIE "PieChartExample"                
  '(                                               
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
))                                                 
```



```racket
(add-chart-sheet                                   
  "Pie3DChart" 'PIE3D "Pie3DChartExample"          
  '(                                               
    ("CAT" "DataSheet" "A1-D1" "DataSheet" "A2-D2")
))                                                 
```


