# analysis-of-ASX200-index-changes
This repo contains SQL and VBA code used to produce visualizations for the investigation of (potential) trading strategies
to make money off changes in the constituents of ASX200 index.

Three types of visualization are produced. The first is an event-study time series plot of cumulative abnormal returns, assuming perfect foresight of the changes: 

![CAR1b](https://user-images.githubusercontent.com/65940824/197307348-b53b997a-8a84-46d1-9822-dad4377b3edf.png)

The calculations to produce this chart are in event_study.sql.

The next type of visualization is a series of scatterplots with categories. One such scatterplot is:

 
The X and Y variables are calculated before each announcement of a change in index membership. The series of scatterplots is used to determine whether patterns of these variables is useful in predicting the announced changes.

The X and Y variables are rankings based on the variables liquidity and average market capitalisation. Liquidity and average market capitalisation are calculated in accordance with Standard and Poors’ index methodology for every ASX stock 10, 20 and 30 trading days before each announcement. These calculations are in create_index_change_factors.sql.

The stored procedure index_change_analysis returns a set of records corresponding to a given announcement date (of which there are four every year) and a given number of trading days prior to the announcement. Each record has fields necessary to produce the scatterplot in Excel.

The stored procedure is called from within VBA in Excel. The relevant code is in index_changes.bas, which imports the result into an ADODB recordset, copies the data into Excel, and produces the scatterplot.


All of the calculation were done using SQL in a SQL Server database

All the scatterplots were created in Excel using VBA to<br>
(1) select a particular index change announcement<br>
(2) select an "offset" in trading days prior to that announcement<br>
(3) extract the necessary data from the SQL Server database<br>
(4) create the resulting scatterplot<br>
<br>
WIP: add additional info re data dictionaries; more comments in code
