SELECT TOP (100) [MessageClass]
               ,[ID]
               ,[ObjectXML]
               ,[LastUpdated]
FROM [BPS].[dbo].[MessageObjects]
WHERE MessageClass = 'MessageClassesPlantRing.B2MML.SyncProductInformationType' and
                 ID like '%HB2530AI70%'
               order by LastUpdated desc