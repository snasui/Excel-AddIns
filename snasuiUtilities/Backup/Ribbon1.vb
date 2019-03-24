Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core


'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("snasuiUtilities.Ribbon1.xml")
    End Function

#Region "Ribbon Callbacks"

    'Public xlApp As Excel.Application
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        'Declare xlApp in module1 as public variable
        xlApp = Globals.ThisAddIn.Application
    End Sub

    Public Sub IRibbon_Click(control As IRibbonControl)
        'Dim wb As Workbook = xlApp.ActiveWorkbook
        'Dim sh As Worksheet = wb.ActiveSheet
        'Dim myRng As Range = sh.Range("A1:A10")
        'Dim i As Long = Module1.TotalVal(myRng)
        Dim ab As AboutBox1, lgin As LoginForm1
        Select Case control.Id
            Case "btnCleanName"
                'MsgBox(i.ToString)
                Call Module2.DeleteNamedRangesError()
            Case "btnCleanStyle"
                Call Module2.DeleteAllNonBuildInStyles()
            Case "btnCleanObjects"
                Call Module2.DeleteAllShapes()
            Case "btnAbout"
                ab = New AboutBox1
                ab.Show()
            Case "btnLogin"
                lgin = New LoginForm1
                lgin.Show()
        End Select
    End Sub

#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
