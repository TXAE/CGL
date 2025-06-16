SELECT TOP (1000) [MessageClass]
      ,[ID]
      ,[ObjectXML]
      ,[LastUpdated]
  FROM [BPS].[dbo].[MessageObjects]
  where id like '%943711%'
  and LastUpdated like '%2025%'