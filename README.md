# analysis-of-ASX200-index-changes
This repo contains SQL and VBA code used to produce visualizations for the investigation of (potential) trading strategies
to make money off changes in the constituents of ASX200 index.

Three types of visualization are produced. The first is an event-study time series plot of cumulative abnormal returns, assuming perfect foresight of the changes: 

![CAR1b](https://user-images.githubusercontent.com/65940824/197307348-b53b997a-8a84-46d1-9822-dad4377b3edf.png)

The calculations to produce this chart are in **event_study.sql**.

The next type of visualization is a series of scatterplots with categories. One such scatterplot is:

![scat1a](https://user-images.githubusercontent.com/65940824/197311539-34b2a393-0a51-4741-aafc-5e343f9594f0.png)

 
The X and Y variables are calculated before each announcement of a change in index membership. The series of scatterplots is used to determine whether patterns of these variables is useful in predicting the announced changes[^1]. These calculations are in **create_index_change_factors.sql**.

The stored procedure **index_change_analysis** returns a set of records corresponding to a given announcement date (of which there are four every year) and a given number of trading days prior to the announcement. Each record has fields necessary to produce the scatterplot in Excel.

The stored procedure is called from within VBA in Excel. The relevant code is in **index_changes.bas**, which imports the result into an ADODB recordset, copies the data into Excel, and produces the scatterplot.

Finally, the following graphic was produced of the cumulative returns of a short-selling trading strategy based on selling the bottom ten stocks in the ASX200 index[^2]:

![CAR2b](https://user-images.githubusercontent.com/65940824/197314706-44bbab05-62c3-4947-8ba3-122a2ca365c3.png)

The computations to produce this chart are in **trading_strategy.sql**.

## List of tables and columns referenced

**ASX200_constituents_pre_change**
Column_name | Type | Nullable
--- | --- | ---
StockID |	int |	no
IndexDate	| datetime2 |	no

**Daily_prices**
Column_name | Type | Nullable
--- | --- | ---
StockID |	int |	no
PriceDate	| datetime2 |	no
Open	| float	| yes
High	| float	| yes
Low	| float	| yes
Close	| float	| yes
AdjOpen	| float	| yes
AdjHigh	| float	| yes
AdjLow	| float	| yes
AdjClose	| float	| yes
Volume	| int	| yes
MarketCap	| float	| yes
Shares	| float	| yes

**indexChangeData**
Column_name | Type | Nullable
--- | --- | ---
Change	nvarchar	yes
StockID |	int |	no
EffDate	| datetime2	| yes
AnnDate	| datetime2	| yes

**SP200Data**
Column_name | Type | Nullable
--- | --- | ---
PriceDate	| datetime2 |	no
AccumIndex	| float	| yes
DailyRet	| float	| yes

**stockaccumindex**
Column_name | Type | Nullable
--- | --- | ---
StockID |	int |	no
PriceDate	| datetime2 |	no
AccumIndex	decimal	yes

**tradingdays**
Column_name | Type | Nullable
--- | --- | ---
PriceDate	| datetime2 |	no
EOM	| smallint	| yes
DateOffset	| int	| yes

[^1]: The X and Y variables are rankings based on the variables liquidity and average market capitalisation. Liquidity and average market capitalisation are calculated in accordance with Standard and Poors’ index methodology for every ASX stock 10, 20 and 30 trading days before each announcement.
[^2]: It was determined that prediction of index removals, but not index additions, could feasibly be attempted.
