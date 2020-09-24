SELECT DISTINCT PermitD.Sku, PermitD.Permit, Count(*) AS Expr1
FROM PermitD
GROUP BY PermitD.Sku, PermitD.Permit
HAVING (((Count(*))>1));
