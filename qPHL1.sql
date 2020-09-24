SELECT Left([PH],2) AS PHL1, ProdHierarchy.Des AS PHNam, ProdHierarchy.Srt AS Srt1, ProdHierarchy.WithOHCur, ProdHierarchy.WithOHHst
FROM ProdHierarchy
WHERE (((ProdHierarchy.Lvl)=1));
