VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menghitung Jumlah Substring dalam String"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Contoh ini akan menunjukkan bahwa string 'is' muncul 2 'kali dalam string 'this is my string'.
'Coding ini menggunakan fungsi 'split' yang hanya 'terdapat mulai Visual Basic 6.0 ke atas.

Private Sub Form_Load()
    myString = "this is my string"
    tempString = Split(myString, "is")
    MsgBox "'is' muncul dalam '" & myString & "' sebanyak " & UBound(tempString) & " kali."
End Sub


