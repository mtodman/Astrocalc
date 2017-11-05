
Public Class frmAstroCalc
    Dim intDiameter As Integer
    Dim intFL As Integer
    Dim dblFR As Double
    Dim dblReducer As Double
    Dim dblBarlow As Double
    Dim dblRFR As Double
    Dim intXPixels As Integer
    Dim intYPixels As Integer
    Dim dblSensorWidth As Double
    Dim dblSensorHeight As Double
    Dim dblPixelWidth As Double
    Dim dblPixelHeight As Double
    Dim dblXFOV As Double
    Dim dblYFOV As Double
    Dim dblArcSecsPerPixelX As Double
    Dim dblArcSecsPerPixelY As Double
    Dim intTotalPixels As Integer


    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Integer, ByVal lpFileName As String) _
        As Integer
    Private Declare Ansi Function WritePrivateProfileString _
           Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
           (ByVal lpApplicationName As String, _
           ByVal lpKeyName As String, ByVal lpString As String, _
           ByVal lpFileName As String) As Integer

    '************************************************************
    '*  Purpose :   To Get Configuration Settings from INI file
    '*
    '*  Inputs  :   strKey(String)          Key for the settings
    '*              strSection(String)      Section name
    '*
    '*  Returns :   Settings
    '*
    '************************************************************
    Public Function GetIniSetting(ByVal strKey As String, Optional ByVal strSection As String = "settings") As String
        Dim strValue As String
        Dim intPos As Integer
        On Error GoTo ErrTrap
        strValue = Space(1024)
        GetPrivateProfileString(strSection, strKey, "NOT_FOUND", strValue, 1024, "[ini file path]")
        Do While InStrRev(strValue, " ") = Len(strValue)
            strValue = Mid(strValue, 1, Len(strValue) - 1)
        Loop
        ' to remove a special chr in the last place
        strValue = Mid(strValue, 1, Len(strValue) - 1)
        GetIniSetting = strValue
ErrTrap:
        If Err.Number <> 0 Then
            Err.Raise(Err.Number, , "Error form Fucntions.GetIniSettings " & Err.Description)
        End If
    End Function

    '*************************************************************** 
    '*  Purpose :   To Set Configuration Settings into  INI file
    '*
    '*  Inputs  :   strKey(String)          Key for the settings
    '*              strValue(String)        Value for the key specified
    '*  Returns :   NA
    '*    '****************************************************************
    Public Sub SetIniSettings(ByVal strKey As String, ByVal strValue As String, Optional ByVal strSection As String = "settings")
        Dim intPos As Integer
        On Error GoTo ErrTrap
        WritePrivateProfileString(strSection, strKey, strValue, "[ini file path]")
ErrTrap:
        If Err.Number <> 0 Then Err.Raise(Err.Number, , "Error form Functions.SetIniSettings " & Err.Description)
    End Sub
    
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        intFL = 1
        intDiameter = 1
        dblFR = 1
        dblReducer = 1
        dblBarlow = 1
        dblRFR = 1
        txtDiameter.Text = 1
        txtFL.Text = 1
        txtFR.Text = 1
        cmbFR.Text = 1
        cmbBarlow.Text = 1
        txtRevisedFL.Text = 1

    End Sub

    Private Sub txtDiameter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDiameter.TextChanged
        intDiameter = txtDiameter.Text
        calculate()
    End Sub

    Private Sub txtFL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFL.TextChanged
        intFL = txtFL.Text
        calculate()
    End Sub

    Private Sub cmbFR_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFR.SelectedIndexChanged
        dblReducer = cmbFR.Text
        calculate()
    End Sub
  
       Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbBarlow.SelectedIndexChanged
        dblBarlow = cmbBarlow.Text
        calculate()
    End Sub

    Private Sub cmbTelescope_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTelescope.SelectedIndexChanged
        Select Case cmbTelescope.Text
            Case "8 inch Meade LX200 series - f/10"
                intFL = 2000
                intDiameter = 200
                txtDiameter.Text = 200
                txtFL.Text = 2000
            Case "10 inch Meade LX200 series - f/10"
                intFL = 2500
                intDiameter = 254
                txtDiameter.Text = 254
                txtFL.Text = 2500
            Case "12 inch Meade LX200 series - f/10"
                intFL = 3048
                intDiameter = 305
                txtDiameter.Text = 305
                txtFL.Text = 3048
            Case "14 inch Meade LX200 series - f/10"
                intFL = 3556
                intDiameter = 356
                txtDiameter.Text = 356
                txtFL.Text = 3556
            Case "Synta ED80"
                intFL = 600
                intDiameter = 80
                txtDiameter.Text = 80
                txtFL.Text = 600
            Case "FSQ106"
                intFL = 530
                intDiameter = 106
                txtDiameter.Text = 106
                txtFL.Text = 530
        End Select
        calculate()
    End Sub

    Private Sub txtXPixels_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtXPixels.TextChanged
        intXPixels = txtXPixels.Text
        calculate()
    End Sub

    Private Sub txtYPixels_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtYPixels.TextChanged
        intYPixels = txtYPixels.Text
        calculate()
    End Sub

    Private Sub txtWidth_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWidth.TextChanged
        dblSensorWidth = txtWidth.Text
        calculate()
    End Sub

    Private Sub txtHeight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHeight.TextChanged
        dblSensorHeight = txtHeight.Text
        calculate()
    End Sub

       Private Sub cmbCamera_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCamera.SelectedIndexChanged
        Select Case cmbCamera.Text
            Case "Meade DSI II"
                intXPixels = 752
                txtXPixels.Text = 752
                intYPixels = 582
                txtYPixels.Text = 582
                dblSensorWidth = 5.59
                txtWidth.Text = 5.59
                dblSensorHeight = 4.68
                txtHeight.Text = 4.68
            Case "Meade DSI III"
                intXPixels = 1360
                txtXPixels.Text = 1360
                intYPixels = 1024
                txtYPixels.Text = 1024
                dblSensorWidth = 10.2
                txtWidth.Text = 10.2
                dblSensorHeight = 8.3
                txtHeight.Text = 8.3
            Case "QHY8"
                intXPixels = 3032
                txtXPixels.Text = 3032
                intYPixels = 2016
                txtYPixels.Text = 2016
                dblSensorWidth = 23.649
                txtWidth.Text = 23.649
                dblSensorHeight = 15.742
                txtHeight.Text = 15.742
            Case "QHY9"
                intXPixels = 3358
                txtXPixels.Text = 3358
                intYPixels = 2536
                txtYPixels.Text = 2536
                dblSensorWidth = 18.133
                txtWidth.Text = 18.133
                dblSensorHeight = 13.694
                txtHeight.Text = 13.694
            Case "SBIG 8300"
                intXPixels = 3326
                txtXPixels.Text = 3326
                intYPixels = 2504
                txtYPixels.Text = 2504
                dblSensorWidth = 17.96
                txtWidth.Text = 17.96
                dblSensorHeight = 13.521
                txtHeight.Text = 13.521
            Case "Canon EOS 400D"
                intXPixels = 3888
                txtXPixels.Text = 3888
                intYPixels = 2592
                txtYPixels.Text = 2592
                dblSensorWidth = 22.2
                txtWidth.Text = 22.2
                dblSensorHeight = 14.8
                txtHeight.Text = 14.8
        End Select
        calculate()
    End Sub

    Private Sub calculate()
        dblFR = intFL / intDiameter * dblReducer * dblBarlow
        txtFR.Text = dblFR
        dblRFR = intFL * dblReducer * dblBarlow
        txtRevisedFL.Text = dblRFR
        dblPixelWidth = dblSensorWidth / intXPixels * 1000
        txtXPixelSize.Text = dblPixelWidth
        dblPixelHeight = dblSensorHeight / intYPixels * 1000
        txtYPixelSize.Text = dblPixelHeight
        dblXFOV = (206265 * (dblSensorWidth / (intDiameter * dblFR))) / 60
        txtXFOV.Text = dblXFOV
        dblYFOV = (206265 * (dblSensorHeight / (intDiameter * dblFR))) / 60
        txtYFOV.Text = dblYFOV
        dblArcSecsPerPixelX = dblXFOV * 60 / intXPixels
        txtArcSecsPerPixelX.Text = dblArcSecsPerPixelX
        dblArcSecsPerPixelY = dblYFOV * 60 / intYPixels
        txtArcSecsPerPixelY.Text = dblArcSecsPerPixelY
        intTotalPixels = intXPixels * intYPixels
        txtTotalPixels.Text = intTotalPixels
    End Sub
End Class
