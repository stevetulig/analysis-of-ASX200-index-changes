--Change='Addition': script applies to index additions
--Change='Addition': script applies to index removals

--create a temporary table to do all calculations
--for each addition or removal, create a space to
--store the previous 31 and the next 30 days of returns data
with additions_cte (StockID, AnnDate, AnnDateOffset) as
(
select i.StockID, i.AnnDate, t.DateOffset
from indexChangeData i inner join tradingdays t
on i.AnnDate=t.PriceDate
where i.Change='Addition'
--where i.Change='Removal'
)
select a.StockID, a.AnnDate, a.AnnDateOffset, t.DateOffset, t.PriceDate,
	cast(Null as float) as AccumIndex, cast(Null as float) as ret, cast(Null as float) as mktRet, cast(Null as float) as AbnormalRet
into #temp1
from additions_cte a join tradingdays t
on t.DateOffset between a.AnnDateOffset-31 and a.AnnDateOffset+30
order by a.StockID, t.PriceDate

--add the total return (accumulation) index for each stock
update t set t.AccumIndex=s.AccumIndex
from #temp1 t left join stockaccumindex s
on t.StockID=s.StockID and t.PriceDate=s.PriceDate

--calculate the daily returns
update t1 set t1.ret=t1.AccumIndex/t2.AccumIndex-1
from #temp1 t1 inner join #temp1 t2
on t1.STockID=t2.StockID
and t1.DateOffset=t2.DateOffset+1

--add the benchmark market returns from the ASX/S&P200
update t set t.mktRet=s.DailyRet
from #temp1 t inner join SP200Data s
on t.PriceDate=s.PriceDate

--Calculate the abnormal returns
update #temp1 set AbnormalRet=ret-mktRet

--Now average the abnormal returns for each trading day offset
select (DateOffset-AnnDateOffset) as TradingDayOffset, avg(AbnormalRet) as AR
into #tempAR
from #temp1
where DateOffset-AnnDateOffset>-31
group by DateOffset-AnnDateOffset
order by DateOffset-AnnDateOffset

--Now calculate the cumulative abnormal return
select t1.TradingDayOffset, sum(t2.AR) as CAR
from #tempAR t1 inner join #tempAR t2
on t2.TradingDayOffset<=t1.TradingDayOffset
group by t1.TradingDayOffset
order by t1.TradingDayOffset

--Now repeat for the index removals
drop table #temp1
drop table #tempAR
