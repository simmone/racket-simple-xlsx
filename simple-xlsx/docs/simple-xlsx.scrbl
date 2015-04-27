#lang scribble/manual

@(require (for-label racket))

@title{simple-xlsx package usage}

@author+email["Chen Xiao" "chenxiao770117@gmail.com"]

simple-xlsx package is a package to read and write .xlsx format file.@linebreak{}
.xlsx file is a open xml format file.

@table-of-contents[]

@section[#:tag "install"]{Install}

raco pkg install simple-xlsx

@section{Read}

read from a .xlsx file.@linebreak{}
you can get a specific cell's value or loop for the whole sheet's rows.

@defmodule[simple-xlsx]
@(require (for-label simple-xlsx))

there is also a complete read and write example on github:@link["https://github.com/simmone/racket-simple-xlsx/blob/master/example/example.rkt"]{includedin the source}.

@defproc[(with-input-from-xlsx-file
              [xlsx_file_path (path-string?)]
              [user-proc (-> xlsx_handler void?)])
            void?]{
  read xlsx's main outer func, all read assosiated action is include in user-proc.
}

@defproc[(load-sheet
           [sheet_name (string?)]
           [xlsx_handler (xlsx_handler)])
           void?]{
  load specified sheet by sheet name.@linebreak{}
  must first called before other func, because any other func is based on specified sheet.
}

