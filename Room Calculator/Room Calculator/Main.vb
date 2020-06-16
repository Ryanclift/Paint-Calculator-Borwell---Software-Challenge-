Public Class frmMain
    Private TotalFloorArea As Decimal 'For a regular floor, (TotalFloorArea = AreaSec1), for an irregular floor, (AreaSec1 + AreaSec2 + AreaSec3 + AreaSec4)
    Private RoofArea As Decimal
    Private RoomVolume As Decimal
    Private RoofVolume As Decimal
    Private TotalVolume As Decimal
    Private ReqPaint As Decimal
    Private RegFloor As Boolean 'Stores floor's type regular (true) or not (false)
    Private Sec1 As Integer 'Stores shape of Sec1 of the floor either as a rectangle = (1) or square = (2)
    Private Sec2 As Integer
    Private Sec3 As Integer
    Private Sec4 As Integer
    Private AreaSec1 As Decimal
    Private AreaSec2 As Decimal
    Private AreaSec3 As Decimal
    Private AreaSec4 As Decimal
    Private Roof As Integer 'Stores selected type of roof/celling (1, 2 or 3)
    Private Ctrl As Control 'Used to hide/show specific UI controls
    Private Ctrl2 As Control

    'Sets initial and default starting values
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RegFloor = True
        rdoReg.Checked = True
        rdoPyr.Checked = True
        Roof = 1
        cboNumSecs.SelectedIndex = 0
        cboNumDoors.SelectedIndex = 0
        cboNumWindows.SelectedIndex = 0
        cboNumCoats.SelectedIndex = 0
        For Each Ctrl In grpRoom.Controls
            If Ctrl.Tag = "Door2" Or Ctrl.Tag = "Window2" Then
                Ctrl.Enabled = False
            End If
        Next
    End Sub

    'Sets type of floor to regular to disable irrelavent UI controls
    Private Sub rdoReg_CheckedChanged(sender As Object, e As EventArgs) Handles rdoReg.CheckedChanged
        RegFloor = True
        For Each Ctrl In grpFloor.Controls
            If Ctrl.Tag = "Irreg" Or Ctrl.Tag = "Irreg2" Or Ctrl.Tag = "Irreg3" Or Ctrl.Tag = "Irreg4" Then
                Ctrl.Enabled = False
            End If
        Next
    End Sub

    'Sets type of floor to not regular to enable corresponding UI controls
    Private Sub rdoIrr_CheckedChanged(sender As Object, e As EventArgs) Handles rdoIrr.CheckedChanged
        RegFloor = False
        For Each Ctrl In grpFloor.Controls
            If Ctrl.Tag = "Irreg" Then
                Ctrl.Enabled = True
            End If
        Next
    End Sub

    'Used to detect floor's shape to enable corresponding UI controls and disable irrelavent UI controls depending on user's inputted selection/choice
    Private Sub FloorShape(rec, squ)
        If rec.checked = True Then
            squ.enabled = False
            squ.checked = False
        ElseIf rec.checked = False Then
            squ.enabled = True
        End If
    End Sub

    'Sets selected roof/celling type to enable UI corresponding controls and disable irrelavent UI controls
    Private Sub rdoPyr_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPyr.CheckedChanged
        If rdoPyr.Checked = True Then
            Roof = 1
            lblWidRoof.Enabled = False
            txtWidRoof.Enabled = False
            lblWidRoofM.Enabled = False
        Else
            lblWidRoof.Enabled = True
            txtWidRoof.Enabled = True
            lblWidRoofM.Enabled = True
        End If
    End Sub

    'Sets selected roof/celling type to enable UI corresponding controls and disable irrelavent UI controls
    Private Sub rdoFla_CheckedChanged(sender As Object, e As EventArgs) Handles rdoFla.CheckedChanged
        If rdoFla.Checked = True Then
            Roof = 2
            For Each Ctrl In grpRoof.Controls
                If Ctrl.Tag = "Roof" Then
                    Ctrl.Enabled = False
                End If
            Next
            lblHeightRoom.Enabled = True
            lblHeightRoofM.Enabled = True
            txtHeightRoom.Enabled = True
        Else
            For Each Ctrl In grpRoof.Controls
                If Ctrl.Tag = "Roof" Then
                    Ctrl.Enabled = True
                End If
            Next
        End If
    End Sub

    'Sets selected roof/celling type to enable UI corresponding controls and disable irrelavent UI controls
    Private Sub rdoTri_CheckedChanged(sender As Object, e As EventArgs) Handles rdoTri.CheckedChanged
        If rdoTri.Checked = True Then
            Roof = 3
            lblLenRoof.Text = "Shortest Length of Roof's Base"
            lblWidRoof.Text = "Longest Width of Roof's Base"
        Else
            lblLenRoof.Text = "Length of Roof's Base"
            lblWidRoof.Text = "Width of Roof's Base"
        End If
    End Sub

    'Allows to enable relavent UI controls and disables irrelavent UI controls depending on the selected value by the user
    Private Sub cboNumSections_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNumSecs.SelectedIndexChanged
        If cboNumSecs.SelectedItem = 1 Then
            For Each Ctrl In grpFloor.Controls
                If Ctrl.Tag = "Irreg2" Or Ctrl.Tag = "Irreg3" Or Ctrl.Tag = "Irreg4" Then
                    Ctrl.Enabled = False
                End If
            Next
        ElseIf cboNumSecs.SelectedItem = 2 Then
            For Each Ctrl In grpFloor.Controls
                If Ctrl.Tag = "Irreg2" Then
                    Ctrl.Enabled = True
                    For Each Ctrl2 In grpFloor.Controls
                        If Ctrl2.Tag = "Irreg3" Or Ctrl2.Tag = "Irreg4" Then
                            Ctrl2.Enabled = False
                        End If
                    Next
                End If
            Next
        ElseIf cboNumSecs.SelectedItem = 3 Then
            For Each Ctrl In grpFloor.Controls
                If Ctrl.Tag = "Irreg3" Or Ctrl.Tag = "Irreg2" Then
                    Ctrl.Enabled = True
                    For Each Ctrl2 In grpFloor.Controls
                        If Ctrl2.Tag = "Irreg4" Then
                            Ctrl2.Enabled = False
                        End If
                    Next
                End If
            Next
        ElseIf cboNumSecs.SelectedItem = 4 Then
            For Each Ctrl In grpFloor.Controls
                If Ctrl.Tag = "Irreg4" Or Ctrl.Tag = "Irreg3" Or Ctrl.Tag = "Irreg2" Then
                    Ctrl.Enabled = True
                End If
            Next
        End If
    End Sub

    'Sets shape of Sec1 of the floor to a square (2) and disables the rectangle check box and its corresponding UI controls if square check box is still checked
    Private Sub chkSquSec1_CheckedChanged(sender As Object, e As EventArgs) Handles chkSquSec1.CheckedChanged
        FloorShape(chkSquSec1, chkRecSec1)
        If chkSquSec1.Checked = True Then
            Sec1 = 2
            lblLenSec1.Text = "Side Length"
            txtWidSec1.Enabled = False
            lblWidSec1M.Enabled = False
            lblWidSec1.Enabled = False
        Else
            lblLenSec1.Text = "Length"
            txtWidSec1.Enabled = True
            lblWidSec1M.Enabled = True
            lblWidSec1.Enabled = True
        End If
    End Sub

    'Sets shape of Sec2 of the floor to a square (2) and disables the rectangle check box and its corresponding UI controls if square check box is still checked
    Private Sub chkSquSec2_CheckedChanged(sender As Object, e As EventArgs) Handles chkSquSec2.CheckedChanged
        FloorShape(chkSquSec2, chkRecSec2)
        If chkSquSec2.Checked = True Then
            Sec2 = 2
            lblLenSec2.Text = "Side Length"
            txtWidSec2.Enabled = False
            lblWidSec2M.Enabled = False
            lblWidSec2.Enabled = False
        Else
            lblLenSec2.Text = "Length"
            txtWidSec2.Enabled = True
            lblWidSec2M.Enabled = True
            lblWidSec2.Enabled = True
        End If
    End Sub

    'Sets shape of Sec3 of the floor to a square (2) and disables the rectangle check box and its corresponding UI controls if square check box is still checked
    Private Sub chkSquSec3_CheckedChanged(sender As Object, e As EventArgs) Handles chkSquSec3.CheckedChanged
        FloorShape(chkSquSec3, chkRecSec3)
        If chkSquSec3.Checked = True Then
            Sec3 = 2
            lblLenSec3.Text = "Side Length"
            txtWidSec3.Enabled = False
            lblWidSec3M.Enabled = False
            lblWidSec3.Enabled = False
        Else
            lblLenSec3.Text = "Length"
            txtWidSec3.Enabled = True
            lblWidSec3M.Enabled = True
            lblWidSec3.Enabled = True
        End If
    End Sub

    'Sets shape of Sec4 of the floor to a square (2) and disables the rectangle check box and its corresponding UI controls if square check box is still checked
    Private Sub chkSquSec4_CheckedChanged(sender As Object, e As EventArgs) Handles chkSquSec4.CheckedChanged
        FloorShape(chkSquSec4, chkRecSec4)
        If chkSquSec4.Checked = True Then
            Sec4 = 2
            lblLenSec4.Text = "Side Length"
            txtWidSec4.Enabled = False
            lblWidSec4M.Enabled = False
            lblWidSec4.Enabled = False
        Else
            lblLenSec4.Text = "Length"
            txtWidSec4.Enabled = True
            lblWidSec4M.Enabled = True
            lblWidSec4.Enabled = True
        End If
    End Sub

    'Sets shape of Sec1 of the floor to a rectangle (1) and disables the square check box if rectangle check box is still checked
    Private Sub chkRecSec1_CheckedChanged(sender As Object, e As EventArgs) Handles chkRecSec1.CheckedChanged
        FloorShape(chkRecSec1, chkSquSec1)
        If chkRecSec1.Checked = True Then
            Sec1 = 1
        End If
    End Sub

    'Sets shape of Sec2 of the floor to a rectangle (1) and disables the square check box if rectangle check box is still checked
    Private Sub chkRecSec2_CheckedChanged(sender As Object, e As EventArgs) Handles chkRecSec2.CheckedChanged
        FloorShape(chkRecSec2, chkSquSec2)
        If chkRecSec2.Checked = True Then
            Sec2 = 1
        End If
    End Sub

    'Sets shape of Sec3 of the floor to a rectangle (1) and disables the square check box if rectangle check box is still checked
    Private Sub chkRecSec3_CheckedChanged(sender As Object, e As EventArgs) Handles chkRecSec3.CheckedChanged
        FloorShape(chkRecSec3, chkSquSec3)
        If chkRecSec3.Checked = True Then
            Sec3 = 1
        End If
    End Sub

    'Sets shape of Sec4 of the floor to a rectangle (1) and disables the square check box if rectangle check box is still checked
    Private Sub chkRecSec4_CheckedChanged(sender As Object, e As EventArgs) Handles chkRecSec4.CheckedChanged
        FloorShape(chkRecSec4, chkSquSec4)
        If chkRecSec4.Checked = True Then
            Sec4 = 1
        End If
    End Sub

    'Processes all inputs entered by the user to output (floor's area, room's volume and amount of paint required to paint the walls) for the user
    Private Sub btnCal_Click(sender As Object, e As EventArgs) Handles btnCal.Click
        Try
            SectionsArea(1, 2)
            If RegFloor = True Then
                TotalFloorArea = AreaSec1
            ElseIf RegFloor = False And cboNumSecs.SelectedItem = 2 Then
                TotalFloorArea = AreaSec1 + AreaSec2
            ElseIf RegFloor = False And cboNumSecs.SelectedItem = 3 Then
                TotalFloorArea = AreaSec1 + AreaSec2 + AreaSec3
            ElseIf RegFloor = False And cboNumSecs.SelectedItem = 4 Then
                TotalFloorArea = AreaSec1 + AreaSec2 + AreaSec3 + AreaSec4
            End If
            txtArea.Text = TotalFloorArea
            RoomVolume = txtHeightRoom.Text * TotalFloorArea
            If Roof = 1 Or Roof = 3 Then
                Select Case Roof
                    Case Roof = 1
                        RoofVolume = 1 / 3 * txtLenRoof.Text * txtHeightRoof.Text
                    Case Roof = 3
                        RoofVolume = 1 / 2 * txtLenRoof.Text * txtHeightRoof.Text
                End Select
                TotalVolume = RoofVolume + RoomVolume
            Else
                TotalVolume = RoomVolume
            End If
            txtVolume.Text = TotalVolume
            ReqPaint = txtLenWall1.Text * txtWidWall1.Text + txtLenWall2.Text * txtWidWall2.Text _
            + txtLenWall3.Text * txtWidWall3.Text + txtLenWall4.Text * txtWidWall4.Text
            If cboNumDoors.SelectedItem = 1 Then
                ReqPaint = ReqPaint - txtLenDoor1.Text * txtWidDoor1.Text
            ElseIf cboNumDoors.SelectedItem = 2 Then
                ReqPaint = ReqPaint - txtLenDoor1.Text * txtWidDoor1.Text + txtLenDoor2.Text * txtWidDoor2.Text
            End If
            If cboNumWindows.SelectedItem = 1 Then
                ReqPaint = ReqPaint - txtLenWin1.Text * txtWidWin1.Text
            ElseIf cboNumDoors.SelectedItem = 2 Then
                ReqPaint = ReqPaint - txtLenWin1.Text * txtWidWin1.Text + txtLenWin2.Text * txtWidWin2.Text
            End If
            Select Case cboNumCoats.SelectedItem
                Case cboNumCoats.SelectedItem = 1
                    ReqPaint = ReqPaint * 0.065
                Case cboNumCoats.SelectedItem = 2
                    ReqPaint = ReqPaint * 0.13
                Case cboNumCoats.SelectedItem = 3
                    ReqPaint = ReqPaint * 0.195
            End Select
            txtReqPaint.Text = ReqPaint
        Catch ex As Exception
            MsgBox("Error, please make sure you have selected both the type of room's floor and roof/celling")
            txtArea.Text = ""
            txtVolume.Text = ""
            txtReqPaint.Text = ""
        End Try
    End Sub

    'Sub procedure used to correctly process and calculate all rectangle/square area values of the 4 different sections/sectors of the floor
    Private Sub SectionsArea(a, b)
        If Sec1 = a Then
            AreaSec1 = txtLenSec1.Text * txtWidSec1.Text
        ElseIf Sec1 = b Then
            AreaSec1 = txtLenSec1.Text ^ 2
        End If
        If Sec2 = a Then
            AreaSec2 = txtLenSec2.Text * txtWidSec2.Text
        ElseIf Sec2 = b Then
            AreaSec2 = txtLenSec2.Text ^ 2
        End If
        If Sec3 = a Then
            AreaSec3 = txtLenSec3.Text * txtWidSec3.Text
        ElseIf Sec3 = b Then
            AreaSec3 = txtLenSec3.Text ^ 2
        End If
        If Sec4 = a Then
            AreaSec4 = txtLenSec4.Text * txtWidSec4.Text
        ElseIf Sec4 = b Then
            AreaSec4 = txtLenSec4.Text ^ 2
        End If
    End Sub

    'Clears content of UI controls to reset them back to their default/initial values
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        For Each txtbox In {txtLenSec1, txtWidSec1, txtLenSec2, txtWidSec2, txtLenSec3, txtWidSec3, txtLenSec4, txtWidSec4, txtLenWall1,
            txtWidWall1, txtLenWall2, txtWidWall2, txtLenWall3, txtWidWall3, txtLenWall4, txtWidWall4, txtLenDoor1,
            txtWidDoor1, txtLenDoor2, txtWidDoor2, txtLenWin1, txtWidWin1, txtLenWin2, txtWidWin2, txtHeightRoof, txtHeightRoom,
            txtLenRoof, txtWidRoof, txtArea, txtVolume, txtReqPaint}
            txtbox.Clear()
        Next
        cboNumSecs.SelectedIndex = 0
        cboNumDoors.SelectedIndex = 0
        cboNumWindows.SelectedIndex = 0
        cboNumCoats.SelectedIndex = 0
        For Each chkbox In {chkRecSec1, chkSquSec1,
            chkRecSec2, chkSquSec2,
            chkRecSec3, chkSquSec3,
            chkRecSec4, chkSquSec4}
            chkbox.Checked = False
        Next
        rdoReg.Checked = True
        rdoPyr.Checked = True
    End Sub

    'Allows to enable relavent UI controls and disables irrelavent UI controls depending on the selected value by the user
    Private Sub cboNumDoors_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNumDoors.SelectedIndexChanged
        If cboNumDoors.SelectedItem = 2 Then
            For Each Ctrl In grpRoom.Controls
                If Ctrl.Tag = "Door2" Then
                    Ctrl.Enabled = True
                End If
            Next
        Else
            For Each Ctrl In grpRoom.Controls
                If Ctrl.Tag = "Door2" Then
                    Ctrl.Enabled = False
                End If
            Next
        End If
    End Sub

    'Allows to enable relavent UI controls and disables irrelavent UI controls depending on the selected value by the user
    Private Sub cboNumWindows_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNumWindows.SelectedIndexChanged
        If cboNumWindows.SelectedItem = 2 Then
            For Each Ctrl In grpRoom.Controls
                If Ctrl.Tag = "Window2" Then
                    Ctrl.Enabled = True
                End If
            Next
        Else
            For Each Ctrl In grpRoom.Controls
                If Ctrl.Tag = "Window2" Then
                    Ctrl.Enabled = False
                End If
            Next
        End If
    End Sub

    'Used to verify all textboxes' content to only allow numbers and decimal points to be entered by the user
    Private Sub txtLenSec1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtLenSec1.KeyPress, txtWidSec1.KeyPress,
            txtLenSec2.KeyPress, txtWidSec2.KeyPress, txtLenSec3.KeyPress, txtWidSec3.KeyPress, txtLenSec4.KeyPress,
            txtWidSec4.KeyPress, txtLenWall1.KeyPress, txtWidWall1.KeyPress, txtLenWall2.KeyPress, txtWidWall2.KeyPress,
            txtLenWall3.KeyPress, txtWidWall3.KeyPress, txtLenWall4.KeyPress, txtWidWall4.KeyPress, txtLenDoor1.KeyPress,
            txtWidDoor1.KeyPress, txtLenDoor2.KeyPress, txtWidDoor2.KeyPress, txtLenWin1.KeyPress, txtWidWin1.KeyPress,
            txtLenWin2.KeyPress, txtWidWin2.KeyPress, txtHeightRoof.KeyPress, txtHeightRoom.KeyPress, txtLenRoof.KeyPress,
            txtWidRoof.KeyPress
        For Each txtbox In {txtLenSec1, txtWidSec1, txtLenSec2, txtWidSec2, txtLenSec3, txtWidSec3, txtLenSec4, txtWidSec4, txtLenWall1,
            txtWidWall1, txtLenWall2, txtWidWall2, txtLenWall3, txtWidWall3, txtLenWall4, txtWidWall4, txtLenDoor1,
            txtWidDoor1, txtLenDoor2, txtWidDoor2, txtLenWin1, txtWidWin1, txtLenWin2, txtWidWin2, txtHeightRoof, txtHeightRoom,
            txtLenRoof, txtWidRoof}
            If Not e.KeyChar = “.” And Not Char.IsNumber(e.KeyChar) Then
                e.Handled = True
            End If
            If e.KeyChar = Chr(&H8) Then
                e.Handled = False
            End If
        Next
    End Sub
End Class