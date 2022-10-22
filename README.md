# analysis-of-ASX200-index-changes
This repo contains SQL and VBA code used in the investigation of (potential) trading strategies
to make money off ASX200 index changes. The results of this analysis are in the document "index changes.pdf"

We start by showing a chart of cumulative abnormal returns similar to an event study:



All of the calculation were done using SQL in a SQL Server database

All the scatterplots were created in Excel using VBA to<br>
(1) select a particular index change announcement<br>
(2) select an "offset" in trading days prior to that announcement<br>
(3) extract the necessary data from the SQL Server database<br>
(4) create the resulting scatterplot<br>
<br>
WIP: add additional info re data dictionaries; more comments in code
