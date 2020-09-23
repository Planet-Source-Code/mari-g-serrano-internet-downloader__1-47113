VERSION 5.00
Begin VB.UserControl BajaArchivos 
   AutoRedraw      =   -1  'True
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   FillStyle       =   0  'Solid
   Picture         =   "BajaArchivos.ctx":0000
   ScaleHeight     =   720
   ScaleWidth      =   720
   ToolboxBitmap   =   "BajaArchivos.ctx":0815
End
Attribute VB_Name = "BajaArchivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''hay q joderse... ¡+ sencillo imposible!16/05/2003 22:15.-MaRiØ
'''sId se usa para Identificar diferentes archivos q se bajan

'Very very easy...



''Public Event Progreso(BytesBajados As Long, BytesTotales As Long, sId As String)
''Public Event Completado(Bytes As Long, sId As String)
''Private colDest As New Collection
''
''Public Sub Download(WWWFile As String, sDestino As String, Optional sId As String = "Id")
''    colDest.Add sDestino, sId
''    UserControl.AsyncRead WWWFile, vbAsyncTypeFile, sId, vbAsyncReadForceUpdate
''End Sub
''
''Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
''
''   Name AsyncProp.Value As colDest.Item(AsyncProp.PropertyName)
''   colDest.Remove AsyncProp.PropertyName
''   RaiseEvent Completado(AsyncProp.BytesRead, AsyncProp.PropertyName)
''End Sub
''
''Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
''RaiseEvent Progreso(AsyncProp.BytesRead, AsyncProp.BytesMax, AsyncProp.PropertyName)
''End Sub
''
''Public Sub CancelarDownload(Optional sId As String = "Id")
''    UserControl.CancelAsyncRead sId
''End Sub
''
''Private Sub UserControl_Resize()
''    UserControl.Height = 720
''    UserControl.Width = 720
''End Sub

Public Event Progress(DownLoadedBytes As Long, TotalBytes As Long, sId As String)
Public Event Completed(Bytes As Long, sId As String)
Private colDest As New Collection

Public Sub Download(sWWWFile As String, sDestination As String, Optional sId As String = "Id")
    colDest.Add sDestination, sId
    UserControl.AsyncRead sWWWFile, vbAsyncTypeFile, sId, vbAsyncReadForceUpdate
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
   Name AsyncProp.Value As colDest.Item(AsyncProp.PropertyName)
   colDest.Remove AsyncProp.PropertyName
   RaiseEvent Completed(AsyncProp.BytesRead, AsyncProp.PropertyName)
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    RaiseEvent Progress(AsyncProp.BytesRead, AsyncProp.BytesMax, AsyncProp.PropertyName)
End Sub

Public Sub CancelDownload(Optional sId As String = "Id")
    UserControl.CancelAsyncRead sId
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 720
    UserControl.Width = 720
End Sub
