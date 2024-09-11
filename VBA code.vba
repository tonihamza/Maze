' ====================
' Module 2: Movement and Game Logic
' ====================
' This code should be placed in a separate standard module.

' Public variables to store state information
Public col As Double  ' Movement direction (column)
Public row As Double  ' Movement direction (row)
Public cell As Range  ' Current position of the player
Public name As Variant  ' Name of the current player image
Public k As Integer  ' Used to toggle character animation
Public m As Integer  ' Horizontal adjustment for character positioning
Public n As Integer  ' Vertical adjustment for character positioning
Public cheie As Boolean  ' Boolean flag to check if the key has been found
Public myShape As Shape  ' Reference to the player’s shape object
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

' ============================
' Subroutine: Up
' Moves the character up if no walls are in the way.
' ============================

Public Sub Up()
    col = 0
    row = -1
    m = 0
    n = -10
    ' Check if the player can move up (no walls)
    If Not Selection.Offset(-1, 0).Interior.Color = RGB(214, 108, 20) And _
       Not Selection.Offset(-1, -1).Interior.Color = RGB(214, 108, 20) Then
       
        ' Alternate between character images for movement animation
        If k = 1 Then
            name = "Picture 36"
            HideOtherImages
            k = 0
        Else
            name = "Picture 20"
            HideOtherImages
            k = 1
        End If
        MoveImage
    End If
End Sub

' ============================
' Subroutine: Down
' Moves the character down if no walls are in the way.
' ============================

Public Sub Down()
    col = 0
    row = 1
    m = 0
    n = 10
    ' Check if the player can move down (no walls)
    If Not Selection.Offset(2, 0).Interior.Color = RGB(214, 108, 20) And _
       Not Selection.Offset(2, -1).Interior.Color = RGB(214, 108, 20) Then
       
        ' Alternate between character images for movement animation
        If k = 1 Then
            name = "Picture 12"
            HideOtherImages
            k = 0
        Else
            name = "Picture 22"
            HideOtherImages
            k = 1
        End If
        MoveImage
    End If
End Sub

' ============================
' Subroutine: Left
' Moves the character left if no walls are in the way.
' ============================

Public Sub Left()
    col = -1
    row = 0
    m = -5
    n = 0
    ' Check if the player can move left (no walls)
    If Not Selection.Offset(0, -2).Interior.Color = RGB(214, 108, 20) And _
       Not Selection.Offset(1, -2).Interior.Color = RGB(214, 108, 20) Then
       
        ' Alternate between character images for movement animation
        If k = 1 Then
            name = "Picture 15"
            HideOtherImages
            k = 0
        Else
            name = "Picture 38"
            HideOtherImages
            k = 1
        End If
        MoveImage
    End If
End Sub

' ============================
' Subroutine: Right
' Moves the character right if no walls are in the way.
' ============================

Public Sub Right()
    col = 1
    row = 0
    m = 10
    n = 0
    ' Check if the player can move right (no walls)
    If Not Selection.Offset(0, 1).Interior.Color = RGB(214, 108, 20) And _
       Not Selection.Offset(1, 1).Interior.Color = RGB(214, 108, 20) Then
       
        ' Alternate between character images for movement animation
        If k = 1 Then
            name = "Picture 19"
            HideOtherImages
            k = 0
        Else
            name = "Picture 23"
            HideOtherImages
            k = 1
        End If
        MoveImage
    End If
End Sub

' ============================
' Subroutine: MoveImage
' Moves the player's image and handles key and exit interaction.
' ============================

Public Sub MoveImage()
    ' Display the current character image (based on movement direction)
    Set myShape1 = ActiveSheet.Shapes(name)
    myShape1.Visible = True
    
    ' Scroll the view to follow the player's movement
    ActiveWindow.ScrollRow = ActiveWindow.ScrollRow + row
    ActiveWindow.ScrollColumn = ActiveWindow.ScrollColumn + col
    
    ' Move the character shape to align with the selected cell
    Set myShape2 = ActiveSheet.Shapes("Group 39")
    myShape2.Left = (Selection.Left) - (myShape2.Width / 2) + m
    myShape2.Top = (Selection.Top) - (myShape2.Height / 2) + n
    
    ' Select the next cell in the direction of movement
    Selection.Offset(row, col).Select
    
    ' Check if the player has found the key at cell BP41
    If Selection.Address = "$BP$41" Then
        cheie = True
        ActiveSheet.Shapes("Graphic 5").Visible = False  ' Hide the key image
    End If
    
    ' Check if the player has reached the exit at cell FI94 and has the key
    If Selection.Address = "$FI$94" And cheie = True Then
        MsgBox ("GG")  ' Display game won message
        ' Disable the arrow keys when the game is won
        Application.OnKey "{UP}"
        Application.OnKey "{DOWN}"
        Application.OnKey "{LEFT}"
        Application.OnKey "{RIGHT}"
    End If
End Sub

' ============================
' Subroutine: HideOtherImages
' Hides all character images except for the current one.
' ============================

Private Sub HideOtherImages()
    ActiveSheet.Shapes("Picture 12").Visible = False
    ActiveSheet.Shapes("Picture 15").Visible = False
    ActiveSheet.Shapes("Picture 36").Visible = False
    ActiveSheet.Shapes("Picture 19").Visible = False
    ActiveSheet.Shapes("Picture 20").Visible = False
    ActiveSheet.Shapes("Picture 22").Visible = False
    ActiveSheet.Shapes("Picture 23").Visible = False
    ActiveSheet.Shapes("Picture 38").Visible = False
End Sub

' ============================
' Subroutine: Start
' Initializes the game, places the player at the start position.
' ============================

Public Sub Start()
    ActiveSheet.Shapes("Graphic 5").Visible = True  ' Show the key at the start
    cheie = False  ' Player starts without the key
    Range("BH94").Activate  ' Set the player's starting position
    
    ' Hide all character images except the initial frame
    ActiveSheet.Shapes("Picture 12").Visible = False
    ActiveSheet.Shapes("Picture 15").Visible = False
    ActiveSheet.Shapes("Picture 36").Visible = False
    ActiveSheet.Shapes("Picture 19").Visible = False
    ActiveSheet.Shapes("Picture 20").Visible = False
    ActiveSheet.Shapes("Picture 22").Visible = True  ' Show the initial image
    ActiveSheet.Shapes("Picture 23").Visible = False
    ActiveSheet.Shapes("Picture 38").Visible = False
    
    k = 1  ' Initialize animation frame state
    ActiveWindow.Zoom = 280  ' Set zoom for better visibility
    ActiveWindow.ScrollRow = 76  ' Scroll to starting row
    ActiveWindow.ScrollColumn = 25  ' Scroll to starting column
    
    ' Position the player's shape at the starting cell
    Set myShape2 = ActiveSheet.Shapes("Group 39")
    myShape2.Left = (Selection.Left) - (myShape2.Width / 2)
    myShape2.Top = (Selection.Top) - (myShape2.Height / 2)
End Sub


