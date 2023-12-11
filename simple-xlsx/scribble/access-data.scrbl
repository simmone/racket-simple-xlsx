#lang scribble/manual

@title{Access Data}

@section{cell/cell-value?}

cell is the XLSX's cell's name: 

First is the column's index identified by the alphabet from "A".

Second is the row's index identified by the number from 1.

Example: Cell in the second row, the third column, so cell is "C2".

cell-value can be these type: string, number, date.

@section{Example Sheet Data}

@codeblock{
(add-data-sheet "sheet1" '(
                            (1 2 3)
                            (4 5 6)
                          ))
}

@section{Skip the with-sheet*}

If not have a mass access on a sheet, you can get/set data a sheet directly, not need in a with-sheet* scope.

Example:

1. Normally get a cell's value: 
   @codeblock{
   (with-sheet (lambda () (get-cell "A1") ...))
   }

2. Direct get a cell's value:
   @codeblock{
     (get-sheet-ref-cell 0 "A1")
   }

There are 3 ways to locate a sheet:

sheet-ref: use sheet index, start from 0.

sheet-name: use sheet name, match exactly.

sheet-*name*: use part of sheet name, locate first matched sheet name.

Example:

get-cell have 3 direct function: @codeblock{get-sheet-ref-cell, get-sheet-name-cell, get-sheet-*name*-cell}

@section{get-cell/set-cell!}

@codeblock{
  get-cell (-> string? cell-value?)

  (check-equals (get-cell "B2") 5)
}

Direct function: @codeblock{get-sheet-ref-cell, get-sheet-name-cell, get-sheet-*name*-cell}

@codeblock{
  set-cell! (-> string? cell-value? void?)

  (set-cell! "C1" 8)
}

Direct function: @codeblock{set-sheet-ref-cell!, set-sheet-name-cell!, set-sheet-*name*-cell!}

@section{get-row/set-row!}

Row's index from 1.

@codeblock{
  get-row (-> natural? (listof cell-value?))

  (check-equal? (get-row 1) '(1 2 3))
}

Direct function: @codeblock{get-sheet-ref-row, get-sheet-name-row, get-sheet-*name*-row}

@codeblock{
  set-row! (-> natural? (listof cell-value?) void?)

  (set-row! 1 '(7 8 9))
}

Direct function: @codeblock{set-sheet-ref-row!, set-sheet-name-row!, set-sheet-*name*-row!}

@section{get-rows/set-rows!}

@codeblock{
  get-rows (-> (listof (listof cell-value?)))

  (check-equal? (get-rows) '((1 2 3) (4 5 6)))
}

Direct function: @codeblock{get-sheet-ref-rows, get-sheet-name-rows, get-sheet-*name*-rows}

@codeblock{
  set-rows! (-> (listof (listof cell-value?)) void?)

  (set-rows! '((1 2 3) (7 8 9)))
}

Direct function: @codeblock{set-sheet-ref-rows!, set-sheet-name-rows!, set-sheet-*name*-rows!}

@section{get-col/set-col!}

Col's index from 1.

@codeblock{
  get-col (-> natural? (listof cell-value?))
  
  (check-equal? (get-col 1) '(1 4))
  (check-equal? (get-col 2) '(2 5))
  (check-equal? (get-col 3) '(3 6))
}

Direct function: @codeblock{get-sheet-ref-col, get-sheet-name-col, get-sheet-*name*-col}

@codeblock{
  set-col! (-> natural? (listof cell-value?) void?)

  (set-col! 1 '(7 8))
}

Direct function: @codeblock{set-sheet-ref-col!, set-sheet-name-col!, set-sheet-*name*-col!}

@section{get-cols/set-cols!}

@codeblock{
  get-cols (-> (listof (listof cell-value?)))
  
  (check-equal? (get-cols) '((1 4) (2 5) (3 6)))
}

Direct function: @codeblock{get-sheet-ref-cols, get-sheet-name-cols, get-sheet-*name*-cols}

@codeblock{
  set-cols! (-> (listof (listof cell-value?)) void?)
  
  (set-cols! '((7 8) (9 0) (1 2)))
}

Direct function: @codeblock{set-sheet-ref-cols!, set-sheet-name-cols!, set-sheet-*name*-cols!}

@section{get-rows-count/get-cols-count}

@codeblock{
  get-rows-count (-> natural?)
  get-cols-count (-> natural?)
  
  (check-equal? (get-rows-count) 2)
  (check-equal? (get-cols-count) 3)
}

Direct function: 
@codeblock{get-sheet-ref-rows-count, get-sheet-name-rows-count, get-sheet-*name*-rows-count}
@codeblock{get-sheet-ref-cols-count, get-sheet-name-cols-count, get-sheet-*name*-cols-count}