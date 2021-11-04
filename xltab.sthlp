{smcl}
{p2colset 4 18 20 2}{...}
{p2col:{bf:xltab} {hline 2}}Create extended tables and tabulations, with exporting opportunities (beta version 0.1){p_end}
{p2colreset}{...}


{marker description}{...}
{title:Description}

{p 4 4 2}
{bf:xlsheet} is a package that mimicks the behaviour of {bf:tab} and {bf:table} to create tabulations. It replaces the dataset with the created tabulation, allowing for exporting the tabulation easily.
{p_end}

{title:Beta version}

{p 4 4 2}
This package is still under further development. Newer versions will be made available as development continues.
{p_end}

{marker syntax}{...}
{title:Syntax}

{p 4}
One-way tabulation
{p_end}
{p 8 18 2}
{cmdab:xltab} rowvar
{ifin}
[{it:{help xltab##weight:weight}}]
[{cmd:,}
{it:oneway_options}]
{p_end}

{p 4}
Two-way and three-way tabulation:
{p_end}
{p 8 18 2}
{cmdab:xltab} [superrowvar] rowvar colvar
{ifin}
[{it:{help xltab##weight:weight}}]
[{cmd:,}
{it:twoway_options}]

{synoptset 21 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:One-way tabulation}
{synopt:{opt share:s}}display column shares{p_end}
{synopt:{opt perc:entage}}display shares as percentage (only when shares is specified){p_end}

{syntab:Two-way and three-way tabulation}
{synopt:{opt row}}display shares for each row{p_end}
{synopt:{opt perc:entage}}display shares as percentage (only when row is specified){p_end}
{synopt:{opt rowtotal}}add an additional column with the total{p_end}
{synopt:{opt index(val)}}create an index number based on the column with the specified value{p_end}

{syntab:General options}
{synopt:{opt stat(stat)}}calculate statistic (e.g.) sum, mean, p10 etc. (same options as {help collapse:collapse}){p_end}
{synopt:{opt var(varname)}}variable to be used for the statistic{p_end}
{synopt:{opt excl:ude(#)}}exclude a particular value{p_end}
{synopt:{opt missing}}display missing values{p_end}
{synopt:{opt cumul:ative}}provide running totals (by row){p_end}
{synopt:{opt zero}}put zero instead of missing when a cell does not have data{p_end}
{synopt:{opt cmd(stcmd)}}run {it:stcmd} after creating the table (but before saving)){p_end}
{synopt:{opt save(filename)}}save the tabulation as a dta file (.dta not needed to be specified){p_end}

{syntab:Exporting to Excel using xlsheet}
{synopt:{opt table(tableopt)}}export tabulation to Excel spreadsheet (use {bf:auto} if no {it:tableopt} specified). The options are the same options as for {help xlsheet:xlsheet importdata}{p_end}
{synopt:{opt graph(graphopt)}}create an Excel graph based on the tabulation (use {bf:auto} if no {it:graphopt} specified). The options are the same options as for {help xlsheet:xlsheet graph}{p_end}

{synoptline}
{p2colreset}{...}
  {marker weight}{...}
{p 4 6 2}
  {opt aweight}s, {opt fweight}s, {opt iweight}s and {opt pweight}s, are allowed. Please note that the {help collapse:collapse} command is used in the backend, so the same restrictions regarding weights apply.
  {p_end}
  
{title:Examples}
	{p 4 4 2}Please note that {bf:xltab} will replace the current dataset with the tabulation created (unless the table() option is specified). Use of {help preserve:preserve} and {help restore:restore} is suggested.{p_end}
	
    {p 4 4 2}One-way tabulations{p_end}
{cmd}
	// Create a sector breakdown of number of firms by year:
	xltab Year
	
	// Show number of jobs by year:
	xltab Year, stat(sum) var(Employment)
	
	// Show number of firms by sector
	xltab Sector if Year == 2017
	
	// Show number of jobs by sector
	xltab Sector if Year == 2017, stat(sum) var(Employment)
	
	// Show share of firms by sector
	xltab Sector if Year == 2017, shares
	
	// Show percentages (e.g. 50) instead of shares (0.50)
	xltab Sector if Year == 2017, shares percentage 
{txt}
	{p 4 4 2}Two-way tabulations{p_end}
{cmd}
	// Show number of firms by sector and year
	// Note that year will be the row variable, and sector the column variable.
	xltab Year Sector
	
	// Show number of jobs by sector and year, and provide a row total
	xltab Year Sector, stat(sum) var(Employment) rowtotal
	
	// Show share of firms by sector and year (share of firms in each sector for each year)
	xltab Year Sector, stat(sum) var(Employment) row
	
	
{txt}{title:Exporting to Excel using {help xlsheet:xlsheet}}
	{p 4 4 2}The {opt table(tableopt)} and {opt graph(graphopt)} options can be used to automatically call the {help xlsheet:xlsheet importdata} and {help xlsheet:xlsheet graph} command.{p_end}
	{p 4 4 2}If these option are specified, the data will not be replaced (so no use of {help preserve:preserve} and {help restore:restore} is needed). {p_end}
	    {p 4 4 2}An example:{p_end}
	
{cmd}
	xlsheet newbook using "spreadsheet.xlsx"
	xlsheet newsheet, name(By sector) title(Breakdown by sector)
	
	xltab Year Sector, stat(mean) var(L) table(title(Share of employment by sector) ) graph( type(column) subtype(stacked) datalabels(value) datalabelformat(0%) )
	xlsheet close
{txt}
	{p 4 4 2}Please note that this code would be equivalent to:{p_end}
{cmd}
	xlsheet newbook using "spreadsheet.xlsx"
	xlsheet newsheet, name(By sector) title(Breakdown by sector)
	
	preserve
		xltab Year Sector, stat(mean) var(L) 
		xlsheet importdata, title(Share of employment by sector)
		xlsheet graph, type(column) subtype(stacked) datalabels(value) datalabelformat(0%)
	restore
	
	xlsheet close
{txt}
	{p 4 4 2}Set the options of {opt table(tableopt)} and {opt graph(graphopt)} to {bf:auto} in case no options are specified:{p_end}
{cmd}
	xlsheet newbook using "spreadsheet.xlsx"
	xlsheet newsheet, name(By sector) title(Breakdown by sector)
	
	xltab Year Sector, stat(mean) var(L) table(auto) graph(auto)
	xlsheet close
{txt}

{title:Similarities to tab and table}

	{p 4 4 2}Please note that the abilities of {bf:xltab} are similar to {bf:tab} and {bf:table}.{p_end}
	
{cmd}
	tab Year Sector, row
	xltab Year Sector, row

	table Year Sector, c(sum Employment)
	xltab Year Sector, stat(sum) var(Employment)
{txt}

{title:Known issues}

{p 4 4 2}
As this is under development there are still a few issues that could come up:
{p_end}

{p 4 4 2}
- Currently the code does not deal well with negative categorical values as column variables. So if your variable has codes 1 = "Small firm", 2 = "Large firm" and -9 = "Don't know" using this as a column variable could lead to issues.
{p_end}

{title:Author}
    {p 4 4 2}Elwyn Davies, World Bank (edavies@worldbank.org){p_end}
