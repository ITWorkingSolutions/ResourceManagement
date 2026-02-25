CREATE VIEW vwDim_Resource AS
SELECT
    r.ResourceID,
    r.PreferredName,
    r.FirstName,
    r.LastName,
    r.SalutationID,
    s.ListItemName AS SalutationName,
    r.GenderID,
    g.ListItemName AS GenderName,
    r.Email,
    r.Phone,
    r.StartDate,
    r.EndDate,
   LOWER(
    CASE
        WHEN (r.StartDate IS NULL OR r.StartDate <= DATE('now','localtime'))
         AND (r.EndDate IS NULL OR r.EndDate >= DATE('now','localtime'))
        THEN 'true'
        ELSE 'false'
    END
   ) AS IsActiveOnDate
FROM tblResource r
LEFT JOIN tblListItem s ON r.SalutationID = s.ListItemID
LEFT JOIN tblListItem g ON r.GenderID = g.ListItemID;