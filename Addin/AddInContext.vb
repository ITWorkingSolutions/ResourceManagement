' ============================================================================================
'  Class: AddInContext
'  Purpose:
'       Holds application-wide state such as the loaded DatabaseConfig.
'
'  Notes:
'       - Assigned once at add-in startup.
'       - Accessed throughout the add-in via AddInContext.Current.
' ============================================================================================
Friend Class AddInContext

  Friend Property Config As DatabaseConfig

  Friend Shared Property Current As AddInContext

End Class
