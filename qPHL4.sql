SELECT Left([PH],2) AS PHL1, Left([PH],4) AS PHL2, Left([PH],7) AS PHL3, ProdHierarchy.PH AS PHL4, ProdHierarchy.Des AS PHQly, ProdHierarchy.Srt AS Srt4, ProdHierarchy.WithOHCur, ProdHierarchy.WithOHHst
FROM ProdHierarchy
WHERE (((ProdHierarchy.Lvl)=4));
