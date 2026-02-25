CREATE VIEW vwDim_List AS
SELECT
  lt.ListItemTypeID   AS ListTypeID,
  lt.ListItemTypeName AS ListTypeName,
  lt.IsSystemType     AS IsSystemType,
  li.ListItemID       AS ItemID,
  li.ListItemName     AS ListItemName
FROM tblListItem li
JOIN tblListItemType lt ON li.ListItemTypeID = lt.ListItemTypeID;