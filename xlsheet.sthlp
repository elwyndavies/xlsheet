{smcl}

{p2colset 4 16 20 2}{...}
{p2col:{bf:xlsheet} {hline 2}}Create Excel spreadsheets with tables and graphs  (beta version 0.1){p_end}
{p2colreset}{...}

{marker description}{...}
{title:Description}

{p 4 4 2}
{bf:xlsheet} is a package that allows for the interactive creation of spreadsheets including native Excel graphs (also known as charts). The package requires Stata 16 and Python (preferably a distribution like Anaconda) to be installed for the graph generating abilities. If this is not the case, the package will use the {bf:putexcel} package instead for backwards compatibility (but no graphs will be created).
{p_end}

{p 4 4 2}
To generate files {bf:xlsheet} uses the xlsxwriter Pyhon package created by John McNamara. (See also {browse "https://xlsxwriter.readthedocs.io/"})
{p_end}

{title:Beta version}

{p 4 4 2}
This package is still under further development. Newer versions will be made available as development continues.
{p_end}

{marker syntax}{...}
{title:Syntax}

{p 4}
Create new Excel file
{p_end}
{p 8 18 2}
{cmdab:xlsheet newbook}
using filename
{p_end}

{p 4}
Create new Excel worksheet
{p_end}
{p 8 18 2}
{cmdab:xlsheet newsheet},
name({it:sheetname})
{p_end}

{p 4}
Add data to Excel worksheet
{p_end}
{p 8 18 2}
{cmdab:xlsheet importdata},
[{cmd:,}
{it:tableoptions}]
{p_end}

{p 4}
Add a graph to the Excel worksheet based on added data
{p_end}
{p 8 18 2}
{cmdab:xlsheet graph},
[{cmd:,}
{it:graphoptions}]
{p_end}

{p 4}
Add text
{p_end}
{p 8 18 2}
{cmdab:xlsheet text},
text({it:text})
{p_end}

{p 4}
Save the file
{p_end}
{p 8 18 2}
{cmdab:xlsheet close}
{p_end}

{marker syntax}{...}
{title:Syntax}


{p 4 4 2}
{bf:xlsheet newsheet} creates a new spreadsheet, with the filename specified. This needs to be run before anything else can be done.
{p_end}

{p 4 4 2}
{bf:xlsheet newbook} creates a new workbook within the spreadsheet. This needs to follow after {bf:xlsheet newsheet}. A spreadsheet can consist of several workbooks.
{p_end}

{synoptset 24 tabbed}{...}
{synopthdr:Options for {bf:xlsheet newbook}}
{synoptline}
{synopt:{opt name(str)}}Name of the worksheet (compulsory){p_end}
{synopt:{opt title(str)}}Header of the worksheet (will be added on top of the spreadsheet){p_end}




{p 4 4 2}
{bf:xlsheet importdata} imports the current dataset and adds it to the current worksheet. The following options apply:
{p_end}

{synoptset 24 tabbed}{...}
{synopthdr:Options for {bf:xlsheet importdata}}
{synoptline}
{synopt:{opt title(str)}}Title of the data (added above the table){p_end}
{synopt:{opt subtitle(str)}}Subtitle of the data (added above the table){p_end}
{synopt:{opt format(xlfmt)}}formatting of the data, e.g.: 0, 0.00, 0%, 0.00%{p_end}


{p 4 4 2}
{bf:xlsheet graph} uses the data added by {bf:xlsheet importdata} and creates an Excel chart based on this. The following options apply:
{p_end}

{synoptset 24 tabbed}{...}
{synopthdr:Options for {bf:xlsheet graph}}
{synoptline}
{synopt:{opt type(type)}}the type of graph: {bf:line, column, bar, area, bar, pie, doughnut, scatter, stock, radar}{p_end}
{synopt:{opt subtype(subtype)}}subtype for this graph: {bf:stacked, percent_stacked} (for area, bar, column, line graphs){p_end}
{synopt:{opt catcols(#)}}specify the number of columns that should be used for categories{p_end}
{synopt:{opt datalabels(value)}}display data labels, options: {bf:value}{p_end}
{synopt:{opt datalabelformat(xlfmt)}}data label format in Excel format, e.g.: 0, 0.00, 0%, 0.00%{p_end}
{synopt:{opt datalabelposition(pos)}}position for data label (for line/scatter: {bf:center, right, left, above, below}; for bar/column: {bf:center, inside_base, inside_end, outside_end}){p_end}
{synopt:{opt title(str)}}title for the graph{p_end}
{synopt:{opt reversex}}reverse the x-axis{p_end}
{synopt:{opt reversey}}reverse the y-axis{p_end}
{synopt:{opt xsize(#)}}width of the chart in inches (standard value: 5){p_end}
{synopt:{opt ysize(#)}}height of the chart in inches (standard value: 3){p_end}
{synopt:{opt xtitle(str)}}title of the x axis{p_end}
{synopt:{opt ytitle(str)}}title of the y axis{p_end}
{synopt:{opt legendposition}}position of the legend: {bf:top, bottom, left, right, overlay_left, overlay_right, none}{p_end}
{synopt:{opt transpose}}use rows as categories, and columns as series{p_end}
{synopt:{opt alsotranspose}}re-run the command, but with the transpose option applied{p_end}

{title:Author}

    {p 4 4 2}Elwyn Davies, World Bank (edavies@worldbank.org){p_end}

