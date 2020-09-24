SELECT Left([PH],2) AS PHL1, Left([PH],4) AS PHL2, Left([PH],7) AS PHL3, ProdHierarchy.Des AS PHQGp, ProdHierarchy.Srt AS Srt3, ProdHierarchy.WithOHCur, ProdHierarchy.WithOHHst
FROM ProdHierarchy
WHERE (((ProdHierarchy.Lvl)=3));
