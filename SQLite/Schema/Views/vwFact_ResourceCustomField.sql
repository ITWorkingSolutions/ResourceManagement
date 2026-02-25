CREATE VIEW vwFact_ResourceCustomField AS
SELECT
    rnv.ResourceID,
    rnv.ResourceListItemID AS FieldID,
    rli.ResourceListItemName AS FieldName,
    rnv.ListItemID AS ValueID,
    li.ListItemName AS ValueName,
    rnv.ResourceListItemValue AS ValueText
FROM tblResourceNameValue rnv
JOIN tblResourceListItem rli
    ON rnv.ResourceListItemID = rli.ResourceListItemID
LEFT JOIN tblListItem li
    ON rnv.ListItemID = li.ListItemID

UNION ALL

SELECT
    rnvl.ResourceID,
    rnvl.ResourceListItemID AS FieldID,
    rli.ResourceListItemName AS FieldName,
    rnvl.ListItemID AS ValueID,
    li.ListItemName AS ValueName,
    NULL AS ValueText
FROM tblResourceNameValueListItem rnvl
JOIN tblResourceListItem rli
    ON rnvl.ResourceListItemID = rli.ResourceListItemID
LEFT JOIN tblListItem li
    ON rnvl.ListItemID = li.ListItemID;