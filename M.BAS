Attribute VB_Name = "M"
Public FileName As String
Public BooksCallByTitle As Boolean
Public TotalIssueBook, RenewalCounter, MaxFineBal, MembershipDuration, MembershipFee, RenewalFees As Integer

Public Sub LoadGlobalVariables()
Dim db As Connection
Dim adoPrimaryRS As Recordset
Set db = New Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & M.FileName & ";"
Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select TotalIssueBooks,RenewalCounter,MaxFineBal, MembershipDuration, MembershipFee, RenewalFees from GlobalVariables ", db, adOpenStatic, adLockOptimistic
''''''''''''''
TotalIssueBook = adoPrimaryRS.Fields(0)
RenewalCounter = adoPrimaryRS.Fields(1)
MaxFineBal = adoPrimaryRS.Fields(2)

MembershipDuration = adoPrimaryRS.Fields(3)
MembershipFee = adoPrimaryRS.Fields(4)
RenewalFees = adoPrimaryRS.Fields(5)
db.Close
End Sub
