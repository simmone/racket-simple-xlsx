#lang scribble/manual
@(require (for-label racket))

@title{simple-xlsx package usage}

@author+email["Chen Xiao" "chenxiao770117@gmail.com"]

simple-xlsx package is a package to read and write .xlsx format file.@linebreak{}
all .xlsx file is a open xml format file.

@section{Install}

raco pkg install simple-xlsx

@section{Read}

read from a .xlsx file.
you can get a specific cell's value or loop for the whole sheet's rows.
