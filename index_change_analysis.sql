USE [Zenith]
GO
/****** Object:  StoredProcedure [dbo].[index_change_analysis_2]    Script Date: 8/10/2022 1:15:43 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/* Extract cross-sectional data at a specified number of days (@days_before)
before the announcement date (@ann_date) of a change in ASX200 composition
This procedure is called from the VBA procedure index_changes.bas which stores the data in an
Excel file and then creates a corresponding scatterplot of the two main variables,
namely liquidity and market capitalisation, and colours  each point (stock) according
to its classification as either (i) an index constituent, (ii) a new addition to the index,
(iii) a removal from the index, or (iv) any other stock
*/
-- =============================================
CREATE PROCEDURE [dbo].[index_change_analysis_2] 
	@ann_date as datetime,
	@days_before as int
AS
BEGIN
	SET NOCOUNT ON;
	DECLARE @prev_date as datetime

-- Set the calculation date @days_before trading days before the announcement date
	select @prev_date=t2.PriceDate
	from tradingdays t1 inner join tradingdays t2
	on t1.DateOffset=t2.DateOffset+@days_before
	and t1.PriceDate=@ann_date;

/*
Get the cross-sectional data (Liquidity and Market Cap),
Store it in a temp table with columns to manipulate
Liquidity (the y-variable) according to the categories above.
Market cap is the x-variable.
This is so that Excel can produce the required scatterplot
*/
	with LIQ_data(StockID, ItemValue, LIQ_rank) as
	(
	select StockID, MedianLiquidity,
		RANK() over (Order by MedianLiquidity DESC) as LIQ_rank
	from index_change_factors where PriceDate=@prev_date
	)
	,
	MCR_data (StockID, MarketCap, MC_rank) as
	(
	select StockID, MarketCap,
		RANK() over (Order by MarketCap DESC) as MC_rank
	from Daily_prices where PriceDate=@prev_date
	)
	select M.StockID, M.MC_rank, L.LIQ_rank,
	Null as constituent_LR, Null as addition_LR, Null as removal_LR, Null as other_LR
	into #temp1
	from LIQ_data L right join MCR_data M
	on L.StockID=M.StockID

-- Set the liquidity for the constituent category stocks
-- Constituent_LR will be a y-variable in the Excel scatterplot
	update #temp1
	set constituent_LR=LIQ_rank
	from #temp1 a, ASX200_constituents_pre_change b
	where a.StockID=b.StockID
	and b.IndexDate=dateadd(d,14,@ann_date)

-- Set the liquidity for the addition category stocks
-- Addition_LR will be a y-variable for a second series in the Excel scatterplot
	update #temp1
	set addition_LR=LIQ_rank
	from #temp1 a, indexChangeData b
	where a.StockID=b.StockID
	and b.AnnDate=@ann_date
	and b.Change='Addition'

-- Set the liquidity for the removal category stocks
-- Removal_LR will be a y-variable for a third series in the Excel scatterplot
	update #temp1
	set removal_LR=LIQ_rank
	from #temp1 a, indexChangeData b
	where a.StockID=b.StockID
	and b.AnnDate=@ann_date
	and b.Change='Removal'

-- Set the liquidity for all other stocks
-- other_LR will be a y-variable for the fourth series in the Excel scatterplot
	update #temp1
	set other_LR=LIQ_rank
	where isnull(constituent_LR,0)+isnull(addition_LR,0)+isnull(removal_LR,0)=0
	
-- limit the x-variable (MC_RANK) to make it easy for
-- Excel to produce nice-looking scatterplots
	delete #temp1
	where MC_RANK>(
	select 100+max(MC_RANK) from #temp1
	where constituent_LR is not null or addition_LR is not null or removal_LR is not null
	)
	
-- limit the y-variable (MC_RANK) to make it easy for
-- Excel to produce nice-looking scatterplots
	delete #temp1
	where LIQ_RANK>(
	select 100+max(LIQ_RANK) from #temp1
	where constituent_LR is not null or addition_LR is not null or removal_LR is not null
	)

-- we no longer need the original liquidity variable
	alter table #temp1
	drop column LIQ_rank

-- output everything to be picked up as a recordset object in Excel	
	select * from #temp1
END
