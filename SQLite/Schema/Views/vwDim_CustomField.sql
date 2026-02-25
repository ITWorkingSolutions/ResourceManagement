CREATE VIEW vwDim_CustomField AS
SELECT
    rli.ResourceListItemID AS FieldID,
    rli.ResourceListItemName AS FieldName,
    rli.ValueType,
    rli.ListItemTypeID,
    lit.ListItemTypeName
FROM tblResourceListItem rli
LEFT JOIN tblListItemType lit
    ON rli.ListItemTypeID = lit.ListItemTypeID;