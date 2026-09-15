SELECT *
FROM t_pm_purorder AS orders
LEFT JOIN t_bd_supplier AS supplier ON orders.fsupplierid = supplier.fid;
