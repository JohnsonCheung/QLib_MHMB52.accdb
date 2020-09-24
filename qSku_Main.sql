SELECT IIf(Left([CdTopaz],3)='UDV','U','M') AS Stm, IIf(Left([CdTopaz],3)='UDV','Diageo','MH') AS Stream, Topaz.CdTopaz, SKU.*, Left([ProdHierarchy],10) AS PHL4, IIf([Unit/AC]=0,0,[Btl/AC]/[Unit/AC]*[Unit/SC]) AS [Btl/SC], IIf(Nz([Unit/AC],0)=0,Null,[Btl/Ac]/[Unit/Ac]) AS [Btl/Unit]
FROM SKU INNER JOIN Topaz ON SKU.Topaz = Topaz.Topaz;
