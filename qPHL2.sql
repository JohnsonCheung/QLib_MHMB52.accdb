SELECT Left([PH],2) AS PHL1, Left([PH],4) AS PHL2, ProdHierarchy.Des AS PHBrd, ProdHierarchy.Srt AS Srt2, ProdHierarchy.WithOHCur, ProdHierarchy.WithOHHst
FROM ProdHierarchy
WHERE (((ProdHierarchy.Lvl)=2));
