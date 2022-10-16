# analysis-of-ASX200-index-changes
This repo contains SQL and VBA code used in the investigation of (potential) trading strategies
to make money off ASX200 index changes. The results of this analysis are in the document "index changes.pdf"

All of the calculation were done using SQL in a SQL Server database

All the scatterplots were created in Excel using VBA to
(1) select a particular index change announcement
(2) select an "offset" in trading days prior to that announcement
(3) extract the necessary data from the SQL Server database
(4) create the resulting scatterplot
