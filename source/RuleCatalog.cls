VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RuleCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'because there is no way to create class instances by name this is a workaround to create
' rule objects by theire names
' for more background info see https://stackoverflow.com/questions/1057670/create-a-new-object-using-the-text-name-of-the-class

Option Explicit

Public FontWhiteList As New rule_fontwhitelist
