Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim wks As Excel.Worksheet, objList As Excel.ListObject
        Dim nativeWorkbook As Excel.Workbook
        Dim xlapp As Excel.Application = Globals.ThisAddIn.Application

        nativeWorkbook = xlapp.ActiveWorkbook

        If nativeWorkbook IsNot Nothing Then
            'Dim vstoWorkbook As Workbook = Globals.Factory.GetVstoObject(nativeWorkbook)

            For Each wks In nativeWorkbook.Worksheets
                For Each objList In wks.ListObjects
                    objList.TableStyle = ""
                    objList.HeaderRowRange.Font.Bold = True
                    objList.Unlist()
                Next objList
            Next wks

        End If
    End Sub
End Class
