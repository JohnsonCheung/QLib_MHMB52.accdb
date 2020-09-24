SELECT SKU.SKU, SKU.SkuDes, SKU.ProdHierarchy
FROM SKU
WHERE (((Left([PRodHierarchy],10)) Not In (Select PH from ProdHierarchy where Lvl=4)));
