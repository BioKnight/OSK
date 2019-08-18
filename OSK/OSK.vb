Imports System.Runtime.InteropServices

Public Class OSK
    <DllImport("user32.dll", EntryPoint:="VkKeyScanExW")>
    Public Shared Function VkKeyScanExW(ByVal ch As Char, ByVal dwhkl As IntPtr) As Short
    End Function

    Private down_Color As Color = Color.Gray
    Private up_Color As Color = Color.Gainsboro
    Private caps As Boolean = False
    Private show_Case As Boolean = False
    Private key As New List(Of Keys)

    Private Function GetKeyFromChar(c As Char) As Keys
        Dim vkKeyCode As Short = VkKeyScanExW(c, InputLanguage.CurrentInputLanguage.Handle)
        Return CType((((vkKeyCode And &HFF00) << 8) Or (vkKeyCode And &HFF)), Keys)
    End Function
    Public Property color_On_Press As Color
        Get
            Return down_Color
        End Get
        Set(value As Color)
            down_Color = value
        End Set
    End Property
    Public Property button_Color As Color
        Get
            Return up_Color
        End Get
        Set(value As Color)
            up_Color = value
        End Set
    End Property
    Public Property buttons_Change_Case As Boolean
        Get
            Return show_Case
        End Get
        Set(value As Boolean)
            show_Case = value
        End Set
    End Property
    Private Sub OSK_Load(sender As Object, e As EventArgs) Handles Me.Load
        For Each cntrl As Control In Me.Controls
            If TypeOf cntrl Is Button Then
                cntrl.BackColor = up_Color
                AddHandler cntrl.MouseDown, AddressOf btn_Down
                AddHandler cntrl.MouseUp, AddressOf btn_Up
                AddHandler cntrl.Click, AddressOf btn_Click
            End If
        Next
        If show_Case Then
            switch_Case()
        End If


    End Sub
    Private Sub btn_Down(sender As Object, e As MouseEventArgs)
        sender.BackColor = down_Color
        Select Case LCase(sender.text)
            Case "tab"
                key.Add(Keys.Tab)
            Case "enter"
                key.Add(Keys.Enter)
            Case "backspace"
                key.Add(Keys.Back)
            Case "ctrl"
                key.Add(Keys.Control)
            Case "alt"
                key.Add(Keys.Alt)
            Case "esc"
                key.Add(Keys.Escape)
            Case "caps lock"
                If show_Case Then
                    switch_Case()
                End If
            Case "shift"
                key.Add(Keys.ShiftKey)
            Case "space"
                key.Add(Keys.Space)
            Case Else
                Dim this_key As Char

                Char.TryParse(sender.text, this_key)
                key.Add(GetKeyFromChar(this_key))
        End Select
        If caps And Not key.Contains(Keys.ShiftKey) Then key.Add(Keys.ShiftKey)
        RaiseEvent Button_Down(key.ToArray)
    End Sub
    Private Sub btn_Up(sender As Object, e As MouseEventArgs)
        If LCase(sender.text) = "caps lock" Then
            If caps Then ' leave it pressed
                sender.BackColor = up_Color
            End If
        Else
            sender.BackColor = up_Color
        End If

        Select Case LCase(sender.text)
            Case "tab"
                key.Add(Keys.Tab)
            Case "enter"
                key.Add(Keys.Enter)
            Case "backspace"
                key.Add(Keys.Back)
            Case "ctrl"
                key.Add(Keys.Control)
            Case "alt"
                key.Add(Keys.Alt)
            Case "esc"
                key.Add(Keys.Escape)
            Case "caps lock"
                If caps Then
                    caps = False
                Else
                    caps = True
                End If
                If show_Case Then
                    switch_Case()
                End If
            Case "shift"
                key.Add(Keys.ShiftKey)
            Case "space"
                key.Add(Keys.Space)
            Case Else
                Dim this_key As Char

                Char.TryParse(sender.text, this_key)
                key.Add(GetKeyFromChar(this_key))
        End Select
        If caps And Not key.Contains(Keys.ShiftKey) Then key.Add(Keys.ShiftKey)
        RaiseEvent Button_Up(key.ToArray)
        key.Clear()
    End Sub
    Private Sub btn_Click(sender As Object, e As MouseEventArgs)
        If caps And Not key.Contains(Keys.ShiftKey) Then key.Add(Keys.ShiftKey)
        RaiseEvent Button_Pressed(key.ToArray)
        Select Case LCase(sender.text)
            Case "tab"
                RaiseEvent Letter_Pressed(vbTab)
            Case "enter"
                'RaiseEvent Letter_Pressed(Keys.Enter)
            Case "backspace"
                'RaiseEvent Letter_Pressed(vbBack)
            Case "ctrl"
            Case "alt"
            Case "esc"
            Case "caps lock"
            Case "shift"
            Case "space"
                RaiseEvent Letter_Pressed(" ")
                RaiseEvent Key_Pressed(New KeyPressEventArgs(" "))
            Case Else
                If caps Then
                    RaiseEvent Letter_Pressed(sender.text)
                    RaiseEvent Key_Pressed(New KeyPressEventArgs(sender.text))
                Else
                    RaiseEvent Letter_Pressed(LCase(sender.text))
                    RaiseEvent Key_Pressed(New KeyPressEventArgs(LCase(sender.text)))
                End If
        End Select

        key.Clear()
    End Sub
    Private Sub switch_Case()
        For Each cntrl As Control In Me.Controls
            Select Case LCase(cntrl.Text)
                Case "tab"

                Case "enter"

                Case "backspace"

                Case "ctrl"

                Case "alt"

                Case "esc"

                Case "caps lock"

                Case "shift"

                Case "space"

                Case Else
                    If TypeOf cntrl Is Button Then
                        If caps Then
                            cntrl.Text = UCase(cntrl.Text)
                        Else
                            cntrl.Text = LCase(cntrl.Text)
                        End If
                    End If
            End Select
        Next
    End Sub

    Public Event Button_Down(Button() As Keys)
    Public Event Button_Up(Button() As Keys)
    Public Event Button_Pressed(Button() As Keys)
    Public Event Letter_Pressed(key As Char)
    Public Event Key_Pressed(key As KeyPressEventArgs)
End Class
