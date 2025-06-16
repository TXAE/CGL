-- https://crglwiki.atlassian.net/wiki/spaces/EMEAMIT/pages/81435757/BPS+Reinickendorf+GENERAL
-- Documentation tab
-- How to retrieve a production schedule from OrderManagement database
-- This is necessary when a production has been planned for an invalid item' 
-- e.g. one that is deleted by Tibco BPS Product Master Data Manager because something is not right in Interspec like: 
-- “Product definition deleted by the PDM because of a bad definition :
-- LMRC -- Total quantity for R100BUT1002 does not match the conching general bill of material. 39.7641 <> 43.36” (from BPSLoggingReport). 
-- After the local Interspec-person fixes the problem' they can either ask the planners to resend the production schedule or it can be done this way:
-- after you got it with the query below:
-- resend it via SMF Raw data publisher (d:\SMFDeploymentPlant\Tools\SmfRawDataPublisher on CEBERL29MP) with Cargill.REI.BPS.ProductionSchedule subject. SMF Raw data publisher does not provide feedback' check \\ceberr29mp.pcg.cargill.com\D$\SMFDeploymentPlant\SMFDaemon to see if your message went through.
SELECT *
FROM [OrderManagement].[dbo].[MessageObjects]
WHERE id LIKE '%65353915%'