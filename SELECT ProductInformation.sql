SELECT TOP (100) [MessageClass]
               ,[ID]
               ,[ObjectXML]
               ,[LastUpdated]
FROM [BPS].[dbo].[MessageObjects]
WHERE MessageClass = 'MessageClassesPlantRing.B2MML.SyncProductInformationType' and
                 ID like '%HC5637AE00%'
               order by LastUpdated desc