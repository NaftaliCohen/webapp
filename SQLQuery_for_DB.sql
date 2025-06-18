

select *
from dbo.consumptions



select *
from dbo.orders_and_receipts_LT



select *
from dbo.InventoryBalance










SELECT Distinct 
          O.ItemCode,
          O.PRICE,
		  O.PurchaseOrderNum ,
          O.PurchaseOrderDate,
          O.GoodsReceiptDate,
		  O.LT_Days,
          C.DocDate  AS ConsumptionDate,
          C.Quantity AS ConsumptionQty,
          I.DocDate  AS InventoryDocDate,   -- << תאריך רישום במלאי
          I.InQty,
          I.OutQty,
		  O.MinLevel,
          O.MaxLevel,
          I.InventoryBalance

   
FROM  dbo.orders_and_receipts_LT  AS O
LEFT  JOIN dbo.consumptions     AS C   
       ON  O.ItemCode  = C.ItemCode
       AND O.DistNumber = C.DistNumber
LEFT  JOIN dbo.InventoryBalance AS I  
          ON I.DocDate  = C.DocDate
		  AND C.ItemCode = I.ItemCode

WHERE O.ItemCode <> '601-0017639'

ORDER BY
      O.ItemCode,
	  I.DocDate ASC ,
      O.PurchaseOrderDate ,
      O.GoodsReceiptDate ,
      C.DocDate ;