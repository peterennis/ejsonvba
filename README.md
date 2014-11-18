# eJsonVBA

The goal of eJsonVBA is to serve as a [JSON](http://www.json.org/) and [EJSON](http://stackoverflow.com/questions/23754969/why-does-meteor-use-ejson-and-not-bson-directly) parser for [VBA](http://msdn.microsoft.com/en-us/library/office/gg264383(v=office.15).aspx). It will allow Microsoft Office applications to communicate with servers that support JSON and servers that support [Meteor](https://www.meteor.com/).

VBA code will use [late binding](http://excelmatters.com/2013/09/23/vba-references-and-early-binding-vs-late-binding/) so that it can be pasted into a module and just work through being more version-independent of the libraries installed.

## What is the Parser Algorithm?

This is synonomous to the question [How do I write my own parser?](http://techblog.procurios.nl/k/n618/news/view/14605/14863/how-do-i-write-my-own-parser-(for-json).html?pageNr=3#thread_339) and credit goes to Patrick van Bergen for his work in that area.

Kudos also to the dormant [vba-json](https://code.google.com/p/vba-json/) project. Initial steps in eJsonVBA development will build on that foundation. It has open [issues](http://code.google.com/p/vba-json/issues/list) and other [problems](http://stackoverflow.com/questions/5773683/excel-vba-parsed-json-object-loop/19359035#19359035) that will be addressed here in a systematic way.

In addition there is the [VB-JSON](http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html) work that may have some roots in vba-json and [JsonBag](http://www.vbforums.com/showthread.php?738845-VB6-JsonBag-Another-JSON-Parser-Generator) - a VB6 implemenation.




