USE [Zenith]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Extract cross-sectional data at a specified point in time
-- before a change in ASX200 composition
-- =============================================
ALTER PROCEDURE [dbo].[index_change_analysis] 
	@change_date as datetime,
	@analysis_date as datetime
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @prev_date_year as int, @prev_date_month as int, @prev_date as datetime

	set @prev_date_year=year(@analysis_date)
	set @prev_date_month=month(@analysis_date);

	with LIQ_data(StockID, ItemValue, LIQ_rank) as
	(
	select StockID, MedianLiquidity,
		RANK() over (Order by MedianLiquidity DESC) as LIQ_rank
	from Liquidity where PriceDate=(
	select PriceDate from tradingdays
	where Year(PriceDate)=@prev_date_year
	and month(PriceDate)=@prev_date_month
	and EOM=1)
	)
	,
	MCR_data (StockID, MC_rank) as
	(
	select StockID, ItemValue as MC_rank
	from MCR where PriceDate=(
	select PriceDate from tradingdays
	where Year(PriceDate)=@prev_date_year
	and month(PriceDate)=@prev_date_month
	and EOM=1)
	)
	select M.StockID, M.MC_rank, L.LIQ_rank,
	Null as constituent_LR, Null as addition_LR, Null as removal_LR, Null as other_LR
	into #temp1
	from LIQ_data L right join MCR_data M
	on L.StockID=M.StockID

	update #temp1
	set constituent_LR=LIQ_rank
	from #temp1 a, ASX200_constituents_pre_change b
	where a.StockID=b.StockID
--	and b.IndexDate=@change_date
	and Year(b.IndexDate)=Year(@change_date)
	and Month(b.IndexDate)=Month(@change_date)

	update #temp1
	set addition_LR=LIQ_rank
	from #temp1 a, indexChangeData b
	where a.StockID=b.StockID
--	and b.AnnDate=@change_date
	and Year(b.AnnDate)=Year(@change_date)
	and Month(b.AnnDate)=Month(@change_date)
	and b.Change='Addition'

	update #temp1
	set removal_LR=LIQ_rank
	from #temp1 a, indexChangeData b
	where a.StockID=b.StockID
--	and b.AnnDate=@change_date
	and Year(b.AnnDate)=Year(@change_date)
	and Month(b.AnnDate)=Month(@change_date)
	and b.Change='Removal'

	update #temp1
	set other_LR=LIQ_rank
	where isnull(constituent_LR,0)+isnull(addition_LR,0)+isnull(removal_LR,0)=0
	
	delete #temp1
	where MC_RANK>(
	select 100+max(MC_RANK) from #temp1
	where constituent_LR is not null or addition_LR is not null or removal_LR is not null
	)
	
	delete #temp1
	where LIQ_RANK>(
	select 100+max(LIQ_RANK) from #temp1
	where constituent_LR is not null or addition_LR is not null or removal_LR is not null
	)

	alter table #temp1
	drop column LIQ_rank
	
	select * from #temp1
END
