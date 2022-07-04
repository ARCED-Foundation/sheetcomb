{smcl}
{* *! version 1.0.0 Mehrab Ali 07jul2022}{...}
{title:Title}

{phang}
{cmd:sheetcomb} {hline 2}
Combine all sheets from an excel file.


{marker syntax}{...}
{title:Syntax}

{p 8 10 2}
{cmd:sheetcomb using} {it:{help filename}}{cmd:,}
[{it:options}]

{synoptset 25}{...}
{synopthdr}
{synoptline}
{* syntab:Options}
{synopt : {opt key(varlist)}}variables for matching observations while merging{p_end}
{synopt : {opt force}}allow string/numeric variable type mismatch without error{p_end}
{synopt : {opt cellra:nge([start][:end])}}excel cell range to load. start and end are specified using standard Excel cell notation,
for example, {cmd: A1}, {cmd: BC2000}, and {cmd: C23}.{p_end}
{synopt : {opt first:row}}treat first row of Excel data as variable names{p_end}
{synopt :{opt all:string}[{cmd:("}{it:{help format:format}}{cmd:")}]}import all Excel data as strings; optionally, specify the numeric display format{p_end}
{synopt :{opt clear}}replace data in memory{p_end}
{synoptline}

{title:Description}

{pstd}
{cmd:sheetcomb} combines all sheets from an excel file. By default it performs {helpb append}. If {opt key()} is provided, it performs {helpb merge} using the specified key variable(s).

{pstd}

{marker remarks}{...}
{title:Remarks}

{pstd}
The GitHub repository for {cmd:sheetcomb} is
{browse "https://github.com/ARCED-Foundation/sheetcomb":here}.
Previous versions may be found there: see the tags.

{marker examples}{...}
{title:Examples}

{pstd}
Combine all the sheets from {cmd:Data.xlsx}
{p_end}{cmd}{...}
{phang2}. sheetcomb using Data.xlsx {p_end}
{txt}{...}

{pstd}
Same as the previous {cmd:sheetcomb} command,
but specifies option {opt key()} to {helpb merge} by key variable(s) {p_end}{cmd}{...}
{phang2}. sheetcomb using Data.xlsx, key(country year){p_end}
{phang2}. sheetcomb using Data.xlsx, key(country year) force cellrange(A6) first allstring {p_end}
{txt}{...}

{marker authors}{...}
{title:Authors}

{pstd}Mehrab Ali{p_end}

{pstd}For questions or suggestions, submit a
{browse "https://github.com/ARCED-Foundation/sheetcomb/issues":GitHub issue}
or e-mail Email: mehrab.ali@arced.foundation.{p_end}

