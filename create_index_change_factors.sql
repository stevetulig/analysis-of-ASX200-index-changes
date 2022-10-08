/*
Script to calculate Liquidity and Average Market Cap of every stock
at various dates before announcements of changes to the ASX200 index.
The calculations are performed 10, 20 and 30 days before each
announcement date.
*/

--the output will be stored in the following table, which will be used 
--by the index_change_analysis stored procedure
create table index_change_factors (StockId int, PriceDate datetime, MedianVal float, AvgeMC float, MedianLiquidity float)
create unique clustered index idx1 on index_change_factors(StockId, PriceDate)

--create a list of dates 10, 20 and 30 trading days prior to each announcement date
--start with the dates 10 days prior
select t2.PriceDate as CalcDate
into #temp1
from tradingdays t1 inner join tradingdays t2
on t1.DateOffset=t2.DateOffset+10
where t1.PriceDate in (select distinct AnnDate from indexChangeData)

--add the dates 20 days prior
insert #temp1 (CalcDate)
select t2.PriceDate
from tradingdays t1 inner join tradingdays t2
on t1.DateOffset=t2.DateOffset+20
where t1.PriceDate in (select distinct AnnDate from indexChangeData)

--add the dates 30 days prior
insert #temp1 (CalcDate)
select t2.PriceDate
from tradingdays t1 inner join tradingdays t2
on t1.DateOffset=t2.DateOffset+30
where t1.PriceDate in (select distinct AnnDate from indexChangeData)

--calculate liquidity and average market cap for each stock
--for each day in the above list
--we will use a cursor to step through each date in #temp1
declare @pdate1 datetime, @pdate0 datetime

declare dateCrsr cursor local forward_only static for
	select DATEADD(m,-6,CalcDate) as date0, CalcDate as date1 from #temp1

open dateCrsr
fetch next from dateCrsr into @pdate0, @pdate1
while @@FETCH_STATUS=0
begin
	insert into index_change_factors(StockId, PriceDate, MedianVal, AvgeMC)
	select distinct StockID, @pdate1,
		Percentile_disc(0.5) within group (ORDER BY cast(Volume as float)*[Close])
		OVER (partition by StockID),
		avg(MarketCap)
		OVER (partition by StockID)
	from Daily_prices
	where PriceDate between @pdate0 and @pdate1
	
	fetch next from dateCrsr into @pdate0, @pdate1

end
close dateCrsr
deallocate dateCrsr

--final calculation: Median daily value traded / average market cap
--Note that market cap is stored in $millions.
update index_change_factors set MedianLiquidity=MedianVal/(AvgeMC*1000000)
where AvgeMC>0

/*
end of script
*/