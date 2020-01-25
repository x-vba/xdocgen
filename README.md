# XDocGen

## Description

[XDocGen](http://x-vba.com/xdocgen) is a utility tool used to generate documentation from VBA source code.
It uses a basic tag syntax which you can place within your source code, and will
generate JSON documentation based off these tags and the procedures in the
Module.

## Usage

XDocGen is written in pure ES6 JavaScript, which means it can be run in the
browser and does not require any external downloads to run. To use it, simply
go to the XDocGen web page and follow the prompts.

## XDocGen Tags

Tags are placed within comments, and take the following format:

'@TagName: this is my tag

Tags can either be placed within a Function or Sub, in which case they will be
Produre-Level Tags, or outside of Functions or Subs, in which case they will be
Module-Level Tags. For more information on the Tag syntax, see the official
documentation.

## Where does my code go?

Since XDocGen is written in pure ES6 JavaScript and has no external
dependencies, your code is never shipped to a server to generate docs. XDocGen
is run purely locally and can be run offline by saving the web page. Additionally,
the source code for XDocGen can be found in the web page in unminified form, 
so you can be sure that your VBA code remains with you.

## XDocGen isn't running?

XDocGen is written in ES6 JavaScript. Some older browsers don't support ES6
JavaScript (notably older versions of Internet Explorer). If XDocGen does not
run in your browser, try using a different browser, such as Chrome, Firefox, or
Safari.

Another common issue is incorrect syntax, or inconsistent @Param Tags compared
to the actual parameter names of the Function or Sub. For more information on Tag
syntax, see the official documentation.

## License

The MIT License (MIT)

Copyright © 2020 Anthony Mancini

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
