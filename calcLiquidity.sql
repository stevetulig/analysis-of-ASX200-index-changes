/*
Query to calculate Median Liquidity as defined in in Standard & Poors' ASX200 Index Methodology,
i.e. median daily value traded divided by average market cap over the previous six months.
As we don't have tick data we define daily value traded as volume X closing price.
*/
use Zenith

declare @pdate1 datetime, @pdate0 datetime
declare dateCrsr cursor local forward_only static for
	with EOM_dates (date0, date1) as
	(
	select e0.PriceDate as t0, e1.PriceDate as t1
	from tradingdays e0 inner join tradingdays e1
	on DATEDIFF(m,e0.PriceDate, e1.PriceDate)=6
	and e0.EOM=1 and e1.EOM=1
	)
	select date0, date1 from EOM_dates

create table Liquidity_SP (StockId int, PriceDate datetime, MedianVal float, AvgeMC float, MedianLiquidity float)

open dateCrsr
fetch next from dateCrsr into @pdate0, @pdate1
while @@FETCH_STATUS=0
begin

	if datepart(year,@pdate1)>=2000
		insert into Liquidity_SP(StockId, PriceDate, MedianVal, AvgeMC)
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
update Liquidity_SP set MedianLiquidity=MedianVal/(AvgeMC*1000000)
where AvgeMC>0

/*
end of script
*/