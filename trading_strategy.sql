--A script for calculating abnormal returns from a trading
--strategy based on shorting the bottom 10 index stocks

--Collect all index constituents as at index change announcement dates
--add the date corresponding to 30 trading days prior
with constituents_cte (StockID, AnnDate, AnnDateOffset) as
(
select a.StockID, t.PriceDate, t.DateOffset-30
from ASX200_constituents_pre_change a inner join tradingdays t
on t.PriceDate=dateadd(d,-14,a.IndexDate)
)
select c.*, t.PriceDate, cast(Null as float) as MC
into #temp1
from constituents_cte c join tradingdays t
on c.AnnDateOffset=t.DateOffset

create unique clustered index idx1 on #temp1 (StockID, PriceDate)

--add the market cap of each constituent
update t1 set MC=d.MarketCap
from #temp1 t1 inner join Daily_prices d
on t1.StockID=d.StockID and t1.PriceDate=d.PriceDate

--some of the market cap data is missing
--we remove all affected observations
delete #temp1 where MC is null

--Identify the bottom 10 stocks by market cap
--add them to a table with space to calculate returns
with sell_list (StockID, AnnDate, AnnDateOffset) as
(
select StockID, AnnDate, AnnDateOffset from
(select StockID, AnnDate, AnnDateOffset,
		RANK() over (partition by AnnDate Order by MC) as MCR
		from #temp1) a
		where a.MCR<=10
)
select s.*, t.DateOffset, t.PriceDate,
	cast(Null as float) as AccumIndex, cast(Null as float) as ret, cast(Null as float) as mktRet, cast(Null as float) as AbnormalRet
into #temp2
from sell_list s join tradingdays t
on t.DateOffset between s.AnnDateOffset-31 and s.AnnDateOffset
order by s.StockID, t.PriceDate

--add the total return (accumulation) index for each stock
update t set t.AccumIndex=s.AccumIndex
from #temp2 t left join stockaccumindex s
on t.StockID=s.StockID and t.PriceDate=s.PriceDate

--calculate the daily returns
update t1 set t1.ret=t1.AccumIndex/t2.AccumIndex-1
from #temp2 t1 inner join #temp2 t2
on t1.STockID=t2.StockID
and t1.DateOffset=t2.DateOffset+1

--add the benchmark market returns from the ASX/S&P200
update t set t.mktRet=s.DailyRet
from #temp2 t inner join SP200Data s
on t.PriceDate=s.PriceDate

--Calculate the abnormal returns
update #temp2 set AbnormalRet=ret-mktRet

--Now average the abnormal returns for each trading day offset
select (DateOffset-AnnDateOffset) as TradingDayOffset, avg(AbnormalRet) as AR
into #tempAR
from #temp2
where DateOffset-AnnDateOffset>-31
group by DateOffset-AnnDateOffset
order by DateOffset-AnnDateOffset

--Now calculate the cumulative abnormal return
select t1.TradingDayOffset, sum(t2.AR) as CAR
from #tempAR t1 inner join #tempAR t2
on t2.TradingDayOffset<=t1.TradingDayOffset
group by t1.TradingDayOffset
order by t1.TradingDayOffset

--clean up
drop table #temp1
drop table #temp2
drop table #tempAR
