-- stolen from https://crglwiki.atlassian.net/wiki/spaces/EMEAMIT/pages/81491894/KB+-+BPS+-+Reinickendorf+-+Conche+Receipt+parameters+are+missing+after+CONCHEN+button+pressing.
SELECT TOP (1000) *
FROM [BPS].[dbo].[MessageObjects]
WHERE ID like 'H%_P' and ID not like 'HB%_P'
and ObjectXML.exist('//*:ProductSegment/*:ID[.="W2K1"]') = 0
and ObjectXML.exist('//*:EquipmentID[.="Reinickendorf"]') = 1 -- unnecessary, just for testing
order by LastUpdated desc