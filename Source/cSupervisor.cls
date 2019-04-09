VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Supervisor
'This class is used as a custom object to hold employee information in data nodes.
'A node is created for each employee storing all of their relevent information (name, ID, title, etc.).

'Attributes
Private pID As String
Private pName As String
Private pTitle As String
Private pSupervisor_ID As String

'Get/Let Methods
Public Property Get ID() As String
ID = pID
End Property

Public Property Let ID(Value As String)
pID = Value
End Property


Public Property Get Name() As String
Name = pName
End Property

Public Property Let Name(Value As String)
pName = Value
End Property

Public Property Get Title() As String
Title = pTitle
End Property

Public Property Let Title(Value As String)
pTitle = Value
End Property

Public Property Get Supervisor_ID() As String
Supervisor_ID = pSupervisor_ID
End Property

Public Property Let Supervisor_ID(Value As String)
pSupervisor_ID = Value
End Property

'Methods
Function empFullInfo() As String
empFullInfo = ID & " " & Name & " " & Title & " " & Supervisor_ID
End Function


