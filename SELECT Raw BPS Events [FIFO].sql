SELECT TOP (1000) [FIFO]
      ,[ReceivedDate]
      ,[EventType]
      ,[EventIndex]
      ,[EventDate]
      ,[EventTime]
      ,[Field01]
      ,[Field02]
      ,[Field03]
      ,[Field04]
      ,[Field05]
      --,[Field06] -- contain nothing
      --,[Field07] -- contain nothing
      ,[Field08]
  FROM [BPS].[BPS].[vw_Raw_BPS_Events]
  WHERE FIFO = 'FIFO_W2K4' and
  Field02 like 'MX-3602%'
  --Field04 = '943070' AND
  --field08 = 'R100BUT1000'
    order by ReceivedDate desc