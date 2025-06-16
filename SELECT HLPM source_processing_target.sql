SELECT TOP (1000) [connection_id]
      ,[source_unit]
      ,[processing_equipment]
      ,[target_unit]
  FROM [BPS].[BPS].[HLPM]
  --WHERE source_unit like '%11525%'