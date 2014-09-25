Module Variantes

    Dim F As New System_Functions

    'User Variants
    Public Login As String = Nothing
    Public Login_Name As String = Nothing

    Public Login_IsAdministrator = False
    Public Login_IsVariantsSPOC = False
    Public Login_IsInfosysEmployee = False
    Public Login_IsReadOnly = False

    Public HSBarPos As Integer = 0
    Public VSBarPos As Integer = 0
    Public Row_Global As New DataGridViewRow
    Public Standart_MessageError As String = "Error: Please Contact System Administrator"

    Public DefaultEmailContact As String = "bullio.m@pg.com"
    Public DraftEmails As Boolean = False

    'File Folders
    Public HTMLPath As String = "C:\P&G\PSS_Variants_DB\html\"
    Public EXCELPath As String = "C:\P&G\PSS_Variants_DB\Excel\"


    Public Project_Name As String = "PSS Variants DB"
    Public Project_Version As String = "Version: 09-19-2014 (" & My.Application.Info.Version.ToString & ")"
    Public CS As String = "Data Source=131.190.74.97,1528;Initial Catalog=Variants_DB;Persist Security Info=True;User ID=developer;Password=hmetal"

    'Notes
    '1)Install R&BI Framework (Contact Gustavo Bolanos)
    'http://bdc-intra529.internal.pg.com/FrameworkUpdater/Exes.aspx?ExeName=Reporting%20%26%20Business%20Innovation%20Framework&UpdateType=0

    '2)Required Access to BW-IFR (SAP Box: GRP)

End Module
