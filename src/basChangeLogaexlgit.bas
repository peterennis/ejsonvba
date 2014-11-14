Attribute VB_Name = "basChangeLogaexlgit"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

'=============================================================================================================================
' Tasks:
' %007 -
' %006 -
' %005 -
' %004 -
' %003 -
' %002 -
' %001 -
' Issues:
' #007 -
' #006 -
' #005 -
' #004 -
' #003 -
' #002 -
' #001 - xla fix in v009 kills workbook properties output
'=============================================================================================================================
'
'
'20140416 - v009 - Add code to deal with xla file
    ' Add simple tasks and issues tracker to change log
'20140415 - v008 - Ignore zzz* object names
    ' Test with THE_SOURCE_FOLDER
'20140415 - v007 - Fix code for late binding
    ' Ref: http://www.pcreview.co.uk/forums/late-binding-vbide-t991467.html
    ' Remove import folder references, clean up code
'20140228 - v006 - OutputListOfExcelProperties
'20140228 - v005 - Bump
    ' VBA project not trusted - set trusted location and developer macros
'20130920 - v004 - Add comment on Sheet1 and ThisWorkbook
'20130711 - v003 - Code moved to aexlgitClass
'20130711 - v002 - Reorganizing, add old xla comment history
    ' 03312007 - v002 - Start of creating aeXL Library
    '                   Remove worksheets. Keep Sheet1 as it is required. Use zzz to put items to sleep.
    ' 04012007 - v005 - Need IsLoaded and module basFindWindow in main application.
    ' 04072007 - v007 - adaept Process Management Menu.
    ' 04192007 - v008 - gblnWorkbookClosing and gblnWorkbookOpening go to application
    '                 - Comment out Workbook_Open and Workbook_Close code so the lib compiles.
    ' 00000000 - v009 - adaept Process Management development
    ' 00000000 - v010 - adaept Process Management development
    ' 00000000 - v011 - adaept Process Management development
    ' 04142009 - v012 - Renamed to aeXLW Library v012.xla
    '                   Remove all functions not used in adaept Process Management.
'20130711 - v001 - Import basGetComputerName, basGlobal, basMenuFunctions, basWorkbook
    ' from "aeXLW Library v012.xla" (20090727)
'20130709 - v000 - Startup version based on example from Ron de Bruin
    ' Ref: http://www.rondebruin.nl/win/s9/win002.htm

