#lang scribble/manual

@title{Top Level: *-xlsx}

All the operations on the xlsx file: write, read, modify, should be placed in top level functions:

write-xlsx: In its scope, add sheets, set data or styles, in the end, generate a new xlsx file.

read-xlsx: In its scope, read from a xlsx file, get data.

read-and-write-xlsx: In its scope, read from a xlsx file, set data or styles, in the end, write back a new file or overlap the original file.

@section{write-xlsx}

@codeblock|{
write-xlsx (-> path-string? procedure? any)
}|

arg1: The output file name.

arg2: user procedure.

@section{read-xlsx}

@codeblock|{
read-xlsx (-> path-string? procedure? any)
}|

arg1: The input file name.

arg2: user procedure.

@section{read-and-write-xlsx}

@codeblock|{
read-and-write-xlsx (-> path-string? path-string? procedure? any)
}|

arg1: The input file name.

arg2: The output file name.

arg3: user procedure.

@section{get-sheet-name-list}

@codeblock|{
get-sheet-name-list (-> (listof string?))
}|

@section{get-sheet-count}

@codeblock|{
get-sheet-count (-> natural?)
}|
