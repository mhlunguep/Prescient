
CREATE PROCEDURE [dbo].[SP_Total_Contracts_Traded_Report]
    @DateFrom DATE,
    @DateTo DATE
AS
BEGIN
    SET NOCOUNT ON;

    -- Temporary table to hold the total contracts traded for each contract
    CREATE TABLE #ContractsTraded (
        FileDate DATE,
        Contract NVARCHAR(255),
        TotalContractsTraded INT
    );

    -- Populate the temporary table with total contracts traded for each contract
    INSERT INTO #ContractsTraded (FileDate, Contract, TotalContractsTraded)
    SELECT 
        [FileDate],
        [Contract],
        SUM([ContractsTraded]) AS TotalContractsTraded
    FROM 
        [dbo].[DailyMTM] 
    WHERE
        [FileDate] BETWEEN @DateFrom AND @DateTo
    GROUP BY 
        [FileDate],
        [Contract];

    -- Calculate the total contracts traded during the specified period
    DECLARE @TotalContracts INT;
    SELECT @TotalContracts = SUM(TotalContractsTraded) FROM #ContractsTraded;

    -- Final result query
    SELECT
        [FileDate],
        [Contract],
        TotalContractsTraded AS [Contracts Traded],
        CAST((100.0 * TotalContractsTraded / NULLIF(@TotalContracts, 0)) AS DECIMAL(10, 2)) AS [% Of Total Contracts Traded]
    FROM
        #ContractsTraded
    WHERE
        TotalContractsTraded > 0
    ORDER BY
        [FileDate], [Contract];

    -- Drop the temporary table
    DROP TABLE #ContractsTraded;
END;
GO


