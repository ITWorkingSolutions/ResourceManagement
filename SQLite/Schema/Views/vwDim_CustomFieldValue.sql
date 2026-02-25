CREATE VIEW vwDim_CustomFieldValue AS
SELECT
    rli.ResourceListItemID AS FieldID,
    li.ListItemID AS ValueID,
    li.ListItemName AS ValueName
FROM tblResourceListItem rli
JOIN tblListItem li
    ON rli.ListItemTypeID = li.ListItemTypeID;