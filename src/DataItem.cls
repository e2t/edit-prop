VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Written in 2014 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Public value As String
Public rawValue As String
Public newValue As String
Public fromAll As Boolean

Public Sub Class_Initialize()
    value = ""
    rawValue = ""
    newValue = ""
    fromAll = False
End Sub
