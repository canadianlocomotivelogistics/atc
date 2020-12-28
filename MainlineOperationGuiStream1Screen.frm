VERSION 4.00
Begin VB.Form MainlineOperationGuiSteam1Screen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   11190
   ClientLeft      =   15
   ClientTop       =   345
   ClientWidth     =   15285
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   Height          =   11595
   Left            =   -45
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   11597.73
   ScaleMode       =   0  'User
   ScaleWidth      =   15405.35
   ShowInTaskbar   =   0   'False
   Top             =   0
   Width           =   15405
   Begin VB.PictureBox PictureBoxLocomotiveCab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      Picture         =   "MainlineOperationGuiStream1Screen.frx":0000
      ScaleHeight     =   783.299
      ScaleMode       =   0  'User
      ScaleWidth      =   1024
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15360
      Begin VB.PictureBox PictureBoxInjectorSteamValveLive 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   30
         Picture         =   "MainlineOperationGuiStream1Screen.frx":240042
         ScaleHeight     =   1575
         ScaleWidth      =   2010
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   3270
         Width           =   2010
      End
      Begin VB.PictureBox PictureBoxInjectorSteamValveExhaust 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   5040
         Picture         =   "MainlineOperationGuiStream1Screen.frx":24A638
         ScaleHeight     =   1545
         ScaleWidth      =   2010
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   3285
         Width           =   2010
      End
      Begin VB.PictureBox PictureBoxDamper 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   6510
         Picture         =   "MainlineOperationGuiStream1Screen.frx":254906
         ScaleHeight     =   1890
         ScaleWidth      =   1335
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   9615
         Width           =   1335
      End
      Begin VB.PictureBox PictureBoxFireBoxDoor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3210
         Left            =   300
         Picture         =   "MainlineOperationGuiStream1Screen.frx":25CD30
         ScaleHeight     =   3210
         ScaleWidth      =   4455
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8205
         Width           =   4455
      End
      Begin VB.PictureBox PictureBoxCylinderCock 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3060
         Left            =   6000
         Picture         =   "MainlineOperationGuiStream1Screen.frx":28B71A
         ScaleHeight     =   3060
         ScaleWidth      =   405
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   8460
         Width           =   405
      End
      Begin VB.PictureBox PictureBoxRegulator 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   5475
         Left            =   10035
         Picture         =   "MainlineOperationGuiStream1Screen.frx":28FA4C
         ScaleHeight     =   5475
         ScaleWidth      =   1695
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   2475
         Width           =   1695
      End
      Begin VB.PictureBox PictureBoxSmallInjectorCompressor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   12180
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2ADF52
         ScaleHeight     =   960
         ScaleWidth      =   645
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   5550
         Width           =   645
      End
      Begin VB.PictureBox PictureBoxAutomaticBrake 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   13230
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2B0094
         ScaleHeight     =   3120
         ScaleWidth      =   1365
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   6045
         Width           =   1365
      End
      Begin VB.PictureBox PictureBoxSand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1320
         Left            =   2790
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2BE116
         ScaleHeight     =   1320
         ScaleWidth      =   165
         TabIndex        =   14
         Top             =   5805
         Width           =   165
      End
      Begin VB.PictureBox PictureBoxInjectorWaterValveLive 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   13260
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2BEDB8
         ScaleHeight     =   420
         ScaleWidth      =   1005
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   10440
         Width           =   1005
      End
      Begin VB.PictureBox PictureBoxInjectorWaterValveExhaust 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   13860
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2C044A
         ScaleHeight     =   420
         ScaleWidth      =   1005
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   10845
         Width           =   1005
      End
      Begin VB.PictureBox PictureBoxBlower 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   120
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2C1ADC
         ScaleHeight     =   1905
         ScaleWidth      =   1785
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   735
         Width           =   1785
      End
      Begin VB.TextBox LabelCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9630
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "MainlineOperationGuiStream1Screen.frx":2CCDB6
         Top             =   10665
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.HScrollBar ScrollTimeAdjustment 
         Height          =   135
         LargeChange     =   10
         Left            =   13800
         Max             =   100
         Min             =   -50
         TabIndex        =   9
         Top             =   495
         Width           =   1335
      End
      Begin VB.CommandButton ButtonVideoSettings 
         Caption         =   "Video Settings"
         Enabled         =   0   'False
         Height          =   255
         Left            =   13800
         TabIndex        =   8
         Top             =   675
         Width           =   1335
      End
      Begin VB.CommandButton ButtonVideo 
         Caption         =   "Video is Off"
         Enabled         =   0   'False
         Height          =   255
         Left            =   13785
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton ButtonHelp 
         Caption         =   "&Help is Off"
         Height          =   255
         Left            =   13785
         TabIndex        =   6
         Top             =   1230
         Width           =   1335
      End
      Begin VB.CommandButton ButtonCaption 
         Caption         =   "&Caption is Off"
         Height          =   255
         Left            =   13785
         TabIndex        =   5
         Top             =   1515
         Width           =   1335
      End
      Begin VB.CommandButton ButtonDetail 
         Caption         =   "&Data is Off"
         Height          =   285
         Left            =   13785
         TabIndex        =   4
         Top             =   1785
         Width           =   1305
      End
      Begin VB.PictureBox PictureBoxPointer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   9656
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2CCDBC
         ScaleHeight     =   345
         ScaleWidth      =   390
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   4860
         Width           =   390
      End
      Begin VB.PictureBox PictureBoxReverser 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   8295
         Picture         =   "MainlineOperationGuiStream1Screen.frx":2CD52E
         ScaleHeight     =   1440
         ScaleWidth      =   4650
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   8670
         Width           =   4650
      End
      Begin VB.CommandButton ButtonClose 
         Caption         =   "&Close"
         Height          =   255
         Left            =   13770
         TabIndex        =   1
         Top             =   2085
         Width           =   1305
      End
   End
   Begin VB.Menu menuCaptureDevice 
      Caption         =   "Capture Device"
      Visible         =   0   'False
      Begin VB.Menu menuCaptureDeviceVideoSource 
         Caption         =   "Video Source"
      End
      Begin VB.Menu menuCaptureDeviceAudioSetting 
         Caption         =   "Audio Setting"
      End
      Begin VB.Menu menuCaptureDeviceVideoFormat 
         Caption         =   "Video Format"
      End
      Begin VB.Menu menuCaptureDeviceVideoCompression 
         Caption         =   "Video Compression"
      End
      Begin VB.Menu menuCaptureDeviceVideoDisplay 
         Caption         =   "Video Display"
      End
   End
End
Attribute VB_Name = "MainlineOperationGuiSteam1Screen"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


















Private Sub ButtonCaption_Click()

If ButtonCaption.Caption = "&Caption is On" Then
    Let ButtonCaption.Caption = "&Caption is Off"
Else
    Let ButtonCaption.Caption = "&Caption is On"
End If

End Sub


Private Sub ButtonClose_Click()

MainlineOperationGuiSteam1Screen.Hide
Unload MainlineOperationGuiSteam1Screen
MainlineOperationGUI.Show vbModeless

End Sub

Private Sub ButtonHelp_Click()

If ButtonHelp.Caption = "&Help is Off" Then
    Let ButtonHelp.Caption = "&Help is On"
Else
    Let ButtonHelp.Caption = "&Help is Off"
End If

End Sub

Private Sub ButtonVideo_Click()

If ButtonVideo.Caption = "Video is Off" Then
    Let ButtonVideo.Caption = "Video is On"
    Let VideoCapture.Visible = True
Else
    Let ButtonVideo.Caption = "Video is Off"
    Let VideoCapture.Visible = False
End If

End Sub

Private Sub ButtonVideoSettings_Click()

With VideoCapture
    If .HasAudio Then
        Let menuCaptureDeviceAudioSetting.Enabled = False
    Else
        Let menuCaptureDeviceAudioSetting.Enabled = True
    End If
    If .HasDlgFormat Then
        Let menuCaptureDeviceVideoFormat.Enabled = True
    Else
        Let menuCaptureDeviceVideoFormat.Enabled = False
    End If
    If .HasDlgDisplay Then
        Let menuCaptureDeviceVideoDisplay.Enabled = True
    Else
        Let menuCaptureDeviceVideoDisplay.Enabled = False
    End If
    If .HasDlgSource Then
        Let menuCaptureDeviceVideoSource.Enabled = True
    Else
        Let menuCaptureDeviceVideoSource.Enabled = False
    End If
End With

MainlineOperationGuiSteam1Screen.PopupMenu menuCaptureDevice

End Sub

Private Sub Form_Load()

'    Left = (Screen.Width - Width) / 2   ' Center form horizontally.
'    Top = (Screen.Height - Height) / 2  ' Center form verti'cally.

Let PictureBoxLocomotiveCab.Picture = LoadPicture(App.Path + "\Graphics\Locomotive Steam1\CabScreen(s1).bmp")

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Checking the Screen Resolution
'
' Every time a new window is opened in Autoamtic Train Control we check the screen size and compare it to the window screen size. If the window cannot be displayed in the current screen size a
' message box is displayed. This allows time for the user to change the screen attributes to correct size.

Do
    If Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY Then
        Let TemporaryResponse = MsgBox("Warning! Automatic Train Control program window called '" & Name & "' requires a minimum of " & Width / Screen.TwipsPerPixelX & " by " & Height / Screen.TwipsPerPixelY & " pixels.  Please change your screen resolution to a larger setting to accomodate this window.", vbRetryCancel + vbExclamation, "ATC - User Error")
        If TemporaryResponse = vbCancel Then
            End
        End If
    End If
Loop While Screen.Width / Screen.TwipsPerPixelX < Width / Screen.TwipsPerPixelX Or Screen.Height / Screen.TwipsPerPixelY < Height / Screen.TwipsPerPixelY
    
' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' End Sub Statement
'
' Ends a procedure or block.
'
' Syntax is in the following format
'
'   End Sub
'
' End Sub Required to end a Sub statement. For Visual Basic in-process OLE server (DLL) considerations and restrictions
' that apply to this topic, see Make OLE DLL File Command. When executed, the End statement resets all module-level
' variables and all static local variables in all modules.  If you need to preserve the value of these variables, use
' the Stop Statement instead.  You can then resume execution while preserving the value of those variables.

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If UnloadMode <> vbFormCode Then
    MsgBox "Please use the Close button. Do not close this window buy eXiting."
    Cancel = True
End If

End Sub






















Private Sub SendCommand()

    Let MainlineOperationGUI!SevenByteD.Text = "0"

' For Next Statement
'
' Repeats a group of statements a specified number of times.
'
'
' The step argument can be either positive or negative.
' The value of the step argument determines loop processing as follows:
'
' Once the loop starts and all statements in the loop have executed, step is added to counter.
' At this point, either the statements in the loop execute again (based on the same test that caused the loop to execute
' initially), or the loop is exited and execution continues with the statement following the Next statement.
' Tip, changing the value of counter while inside a loop can make it more difficult to read and debug your code.
'   The Exit For can only be used within a For Each...Next or For...Next control structure to provide an alternate way to exit.
'   Any number of Exit For statements may be placed anywhere in the loop.
'   The Exit For is often used with the evaluation of some condition (for example, If...Then), and transfers control to the statement immediately following Next.
'   You can nest For...Next loops by placing one For...Next loop within another.
'   Give each loop a unique variable name as its counter.
'
' My Notes:
'
' For Next statement is used to set up a loop for each of the bits in a bit. I'm trying to calculate the error byte; hence,
' I need to look at each byte of the packet. Eight bits to a byte so...

    For X = 1 To 8

' Temporary Counter

' My Notes:
'
' I needed to use a temporary counter to add up all the bits. In each one of the bytes to be sent to the communication port,
' i examine the bits to see if it is one or zero. At the end of this routine, it is used to calculate the error btye. This eror byte is needed to conplete the packet.

    Let temp = 0

' My Notes:
'
' For each one of these 'if statements', we are checking to see if the byte should be sent to the communication port.
' For example, the first byte is the first byte in the locomotive address. The second byte is the second of  the locomotives
' address; which may not always be needed. There for the check is omitted.
'   Once inside the first 'if statment' we preform another 'if statement'. This statement is used to determine if the
'   bit of the byte is equal to one or zero. We are counting the number of one bits to determin the rror code.
'   If the bit is equal to one, then our temporary vaiiable is incremented by one.

    If MainlineOperationGUI!OneByteD.Text <> "" Then
        If MainlineOperationGUI!OneByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!OneByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!TwoByteD.Text <> "" Then
        If MainlineOperationGUI!TwoByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!TwoByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!ThreeByteD.Text <> "" Then
        If MainlineOperationGUI!ThreeByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!ThreeByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!FourByteD.Text <> "" Then
        If MainlineOperationGUI!FourByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!FourByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!FiveByteD.Text <> "" Then
        If MainlineOperationGUI!FiveByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!FiveByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    If MainlineOperationGUI!SixByteD.Text <> "" Then
        If MainlineOperationGUI!SixByteD.Text <> "   " Then
            If Mid$(MainlineOperationGUI!SixByteB.Text, X, 1) = "1" Then
                Let temp = temp + 1
            End If
        End If
    End If
    
' Which Bit?
'
' My Notes:
'
' Since our fornext loop starts at a value of one and continues throu to value of eight, the value of the bit we are
' checking on is placed into a temporary spot. When calculating the error byte we need tuen on the appropriate bit.
        
    If X = 1 Then bitvalue = 128
    If X = 2 Then bitvalue = 64
    If X = 3 Then bitvalue = 32
    If X = 4 Then bitvalue = 16
    If X = 5 Then bitvalue = 8
    If X = 6 Then bitvalue = 4
    If X = 7 Then bitvalue = 2
    If X = 8 Then bitvalue = 1
    
' My Notes:
'
' The last step of the loop is to find out if the total number of ones, is even or odd. This is used in calculating the
' error byte. On the first loop, x =1, and the bitvalue = 128 (most significant bit) and there for if the number of ones,
' is odd then the error bit will be one. This is the 'exclusive or' operation or 'xor'.

    If Int(temp / 2) <> (temp / 2) Then MainlineOperationGUI!SevenByteD.Text = Val(MainlineOperationGUI!SevenByteD.Text) + bitvalue

' My Notes:
'
' This is where we need to return to the top of the 'for next' loop. Again the loop is preformed eight times, once for
' each bit inthe varible.
            
Next X
  

    
' Communication Section
'
' Now that the seventh byte has been calculated, we can proceed to sending the command to the communication port. THis is
' done like any other command set to the communication port.
'
' Before setting the communication port, I used this let statement to set the visual status on the screen. Nost of the
' Screen contain this lable to help notify the user waht is happening with the program.
'
' Let Statements
'
' Two Visual Basic statements are used in combination with the assignment operator (=).
' The Let statement, although usually implicit, is used for assigning values.
' The Set statement, which must always be explicit, is used for assigning object references.
' If you use Let instead of Set when assigning an object reference, you will generally end up assigning the value of the object's default property.
' Attempting to use the resulting variable as an object reference will usually result in an error, such as  Error 424 Object required.

Let MainlineOperationGUI!LocomotiveCommunicationStatus.Caption = "Status: Sending Command"

' As well, I initially set the 'commandcontrol' string to the North Coast Engineering command for sending a command to the
' decoder. The following format is used in sending a packet:
'       's cxx yy yy..'
'   where 's' repersent the command to send a packet
'   where 'c' represent the nottation of number of times to repeat this packet.
'   where 'xx' is the number of times to send this packet in hexidecimal. I've hardcoded this to four.
'   where 'yy' is the data to be sent to the command station, and repeated as often as necessary.
' The last hexidecimal should be the error byte.

        Let CommandControl = "q"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
        If MainlineOperationGUI!OneByteD.Text <> "" Then
            If MainlineOperationGUI!OneByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!OneByteH.Text
            End If
        End If
        If MainlineOperationGUI!TwoByteD.Text <> "" Then
            If MainlineOperationGUI!TwoByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!TwoByteH.Text
            End If
        End If
        If MainlineOperationGUI!ThreeByteD.Text <> "" Then
            If MainlineOperationGUI!ThreeByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!ThreeByteH.Text
            End If
        End If
        If MainlineOperationGUI!FourByteD.Text <> "" Then
            If MainlineOperationGUI!FourByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!FourByteH.Text
            End If
        End If
        If MainlineOperationGUI!FiveByteD.Text <> "" Then
            If MainlineOperationGUI!FiveByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!FiveByteH.Text
            End If
        End If
        If MainlineOperationGUI!SixByteD.Text <> "" Then
            If MainlineOperationGUI!SixByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!SixByteH.Text
            End If
        End If

        If MainlineOperationGUI!SevenByteD.Text <> "" Then
            If MainlineOperationGUI!SevenByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!SevenByteH.Text
            End If
        End If
            
' We finish the string by adding a carriage return to it. The command station will then recognize the command when sent.
            
            Let CommandControl = CommandControl + Chr$(13)
            
' Start Sending the information
'
' The first order of business before sending the command to the communication port is to add the command string to the
' communication window. This communication window is located the the Automatic Train Control Form, and controls which
' controls all the characters going in and out of the communication port.
' The following statement, i believe, set the cursor to the end of the new text being ddisplayed in the communication window.

    Let MainScreen.CommunicationWindow.Text = MainScreen.CommunicationWindow.Text + CommandControl + Chr$(10)
    Let MainScreen.CommunicationWindow.SelStart = Len(MainScreen.CommunicationWindow.Text)

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

    MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent paramters as per National Model Railroader Association specification.
    
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

    While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
        Let temp = DoEvents
    Wend
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...
    
' Communication Section
'
' Now that the seventh byte has been calculated, we can proceed to sending the command to the communication port. THis is
' done like any other command set to the communication port.
'
' Before setting the communication port, I used this let statement to set the visual status on the screen. Nost of the
' Screen contain this lable to help notify the user waht is happening with the program.
'
' Let Statements
'
' Two Visual Basic statements are used in combination with the assignment operator (=).
' The Let statement, although usually implicit, is used for assigning values.
' The Set statement, which must always be explicit, is used for assigning object references.
' If you use Let instead of Set when assigning an object reference, you will generally end up assigning the value of the object's default property.
' Attempting to use the resulting variable as an object reference will usually result in an error, such as  Error 424 Object required.

Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Sending Command"

' As well, I initially set the 'commandcontrol' string to the North Coast Engineering command for sending a command to the
' decoder. The following format is used in sending a packet:
'       's cxx yy yy..'
'   where 's' repersent the command to send a packet
'   where 'c' represent the nottation of number of times to repeat this packet.
'   where 'xx' is the number of times to send this packet in hexidecimal. I've hardcoded this to four.
'   where 'yy' is the data to be sent to the command station, and repeated as often as necessary.
' The last hexidecimal should be the error byte.

If MainScreen!checkboxdequeuepacket.Value = vbChecked Then

        Let CommandControl = "d"
            
' If I am suppose to send the first byte of data (does not contain a null string then add the first byte to the
' 'commandcontrol' string. When the data base is updated, it night be necessary to change the null parameter of the 'if statement'.
            
        If MainlineOperationGUI!OneByteD.Text <> "" Then
            If MainlineOperationGUI!OneByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!OneByteH.Text
            End If
        End If
        If MainlineOperationGUI!TwoByteD.Text <> "" Then
            If MainlineOperationGUI!TwoByteD.Text <> "   " Then
                Let CommandControl = CommandControl + " " + MainlineOperationGUI!TwoByteH.Text
            End If
        End If
        'If mainlineoperationGUI!ThreeByteD.Text <> "" Then
        '    If mainlineoperationGUI!ThreeByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!ThreeByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!FourByteD.Text <> "" Then
        '    If mainlineoperationGUI!FourByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!FourByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!FiveByteD.Text <> "" Then
        '    If mainlineoperationGUI!FiveByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!FiveByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!SixByteD.Text <> "" Then
        '    If mainlineoperationGUI!SixByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!SixByteH.Text
        '    End If
        'End If
        'If mainlineoperationGUI!SevenByteD.Text <> "" Then
        '    If mainlineoperationGUI!SevenByteD.Text <> "   " Then
        '        Let CommandControl = CommandControl + " " + mainlineoperationGUI!SevenByteH.Text
        '    End If
        'End If
            
' We finish the string by adding a carriage return to it. The command station will then recognize the command when sent.
            
            Let CommandControl = CommandControl + Chr$(13)
            
' Start Sending the information
'
' The first order of business before sending the command to the communication port is to add the command string to the
' communication window. This communication window is located the the Automatic Train Control Form, and controls which
' controls all the characters going in and out of the communication port.
' The following statement, i believe, set the cursor to the end of the new text being ddisplayed in the communication window.

    Let MainScreen.CommunicationWindow.Text = MainScreen.CommunicationWindow.Text + CommandControl + Chr$(10)
    Let MainScreen.CommunicationWindow.SelStart = Len(MainScreen.CommunicationWindow.Text)

' Spock to Enterprise
'
' Everything is set, not send the commandcontrol to the Communication port. Please note that other parameters have already
' set in the Auotmatic Train Control Form, with the communication object.

    MainScreen.MSComm1.Output = CommandControl
    
' Just so the user knows, I an setting the communication status label, visible on the current form, to 'clear'. This lets
' user know that the command has been send. This does not mean that the command has been recieved by the locomotive, or is
' sent paramters as per National Model Railroader Association specification.
    
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Command Sent"

' Waiting for a response
'
' I'm waiting for a responce for the command station. There are some bugs with this method of confirming the activity
' of the command station, but its the only one implememnted so far. Once the proplems in the on_comm event are smoothed out,
' it might chnge. For now, it creates a method of waiting for te Command Control before continuing.

    While Right$(MainScreen.CommunicationWindow.Text, 9) <> "COMMAND: "
        Let temp = DoEvents
    Wend
    Let MainlineOperationGUI.LocomotiveCommunicationStatus.Caption = "Status: Clear"
    
' Now that the Locomotive Communication Window has be updated...

End If

' =========================================================================================================================
' Automatic Addtion to Comments line
'


End Sub

Public Sub SetLocomotiveNumber()

If MainlineOperationGUI!ShortAdDress.Value = unvbChecked Then
    Let MainlineOperationGUI!OneByteD.Text = Int(Val(MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text) / 256)
    Let MainlineOperationGUI!TwoByteD.Text = Val(MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text) - (Val(MainlineOperationGUI!OneByteD.Text) * 256)
    Let MainlineOperationGUI!OneByteD.Text = Val(MainlineOperationGUI!OneByteD.Text) + 128 + 64
    Let MainlineOperationGUI!ConsistControlComment.Text = "Loco " + MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text + "; "
End If

If MainlineOperationGUI!ShortAdDress.Value = vbChecked Then
    Let MainlineOperationGUI!OneByteD.Text = Int(Val(MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text))
    Let MainlineOperationGUI!TwoByteD.Text = ""
    Let MainlineOperationGUI!ConsistControlComment.Text = "Consist " + MainlineOperationGUI!ConsistControlMacroLocomotiveNumber.Text + "; "
End If

End Sub

Private Sub SetFunction01234()

Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Function "

Let temporarybyte = 128

If MainlineOperationGUI!ConsistControlFunction0.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 16
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "0 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "0 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction1.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 1
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "1 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "1 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction2.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 2
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "2 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "2 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction3.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 4
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "3 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "3 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction4.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 8
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "4 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "4 Off;"
End If

Let MainlineOperationGUI!ThreeByteD.Text = temporarybyte
Let MainlineOperationGUI!FourByteD.Text = ""
Let MainlineOperationGUI!FiveByteD.Text = ""
Let MainlineOperationGUI!SixByteD.Text = ""

End Sub

Private Sub SetFunction5678()

Let temporarybyte = 128 + 32

If MainlineOperationGUI!ConsistControlFunction5.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 1
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "5 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "5 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction6.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 2
    Let MainlineOperationGUI.ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "6 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "6 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction7.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 4
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "7 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "7 Off;"
End If

If MainlineOperationGUI!ConsistControlFunction8.Value = vbChecked Then
    Let temporarybyte = temporarybyte + 8
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "8 On; "
Else
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "8 Off;"
End If

Let MainlineOperationGUI!ThreeByteD.Text = temporarybyte
Let MainlineOperationGUI!FourByteD.Text = ""
Let MainlineOperationGUI!FiveByteD.Text = ""
Let MainlineOperationGUI!SixByteD.Text = ""

End Sub

Private Sub SetChangeCV()
    
        Let TemporaryByteOne = 0
        Let TemporaryByteTwo = Val(ConsistControlCV.Text) - 1
        
        If TemporaryByteTwo / 512 >= 1 Then
            Let TemporaryByteOne = TemporaryByteOne + 2
            Let TemporaryByteTwo = TemporaryByteTwo - 512
        End If
        
        If TemporaryByteTwo / 256 >= 1 Then
            Let TemporaryByteOne = TemporaryByteOne + 1
            Let TemporaryByteTwo = TemporaryByteTwo - 256
        End If
        
        Let TemporaryByteOne = TemporaryByteOne + 128
        Let TemporaryByteOne = TemporaryByteOne + 64
        Let TemporaryByteOne = TemporaryByteOne + 32
        
       If MainlineOperationGUI!ConsistControlCVRead = vbChecked Then
              Let TemporaryByteOne = TemporaryByteOne + 4
        Else
            Let TemporaryByteOne = TemporaryByteOne + 8 + 4
        End If
        
    Let MainlineOperationGUI!ThreeByteD.Text = TemporaryByteOne
    Let MainlineOperationGUI!FourByteD.Text = TemporaryByteTwo
    Let MainlineOperationGUI!FiveByteD.Text = Val(ConsistControlCVValue.Text)
    Let MainlineOperationGUI!SixByteD.Text = ""
    
 Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + _
     "Change CV" + MainlineOperationGUI!ConsistControlCV.Text + " to " + MainlineOperationGUI!ConsistControlCVValue.Text

  
End Sub

Private Sub SetSpeed()

Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Speed "

If MainlineOperationGUI!ConsistControlSpeed128.Value = vbChecked Then
    ' This routine assembles the byte for speed step mode 128
    Let Temporary = Val(MainlineOperationGUI!ConsistControlSpeed.Value)
    Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + Str$(Temporary) + " of 128 "
    
    If MainlineOperationGUI!ConsistControlDirectionF.Value = vbChecked Then
        Temporary = Temporary + 128 ' add forward direction
        Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Forward"
    Else
        Let MainlineOperationGUI!ConsistControlComment.Text = MainlineOperationGUI!ConsistControlComment.Text + "Reverse"
    End If
    
    Let MainlineOperationGUI!ThreeByteD.Text = 63
    Let MainlineOperationGUI!FourByteD.Text = Temporary
    Let MainlineOperationGUI!FiveByteD.Text = ""
    Let MainlineOperationGUI!SixByteD.Text = ""
Else
    Let Temporary = 64
    If MainlineOperationGUI!ConsistControlDirectionF.Value = vbChecked Then
            Let Temporary = Temporary + 32 ' add forward direction
    End If
    
   If MainlineOperationGUI!ConsistControlSpeed28.Value = vbChecked Then
        'This routine assenmles the byte for speed step mode 28
        Let temp1 = Val(MainlineOperationGUI!ConsistControlSpeed.Value) ' adds the speed
        Let temp2 = temp1 Mod 2
        Let newspeedvalue = Int(temp1 / 2)
        Let Temporary = Temporary + newspeedvalue
        If temp2 = 1 Then Let Temporary = Temporary + 16
        Let MainlineOperationGUI!ThreeByteD.Text = Temporary
        Let MainlineOperationGUI!FourByteD.Text = ""
        Let MainlineOperationGUI!FiveByteD.Text = ""
        Let MainlineOperationGUI!SixByteD.Text = ""
    Else
        ' This routing assembles the byte for speed step mode 14
        
        Let Temporary = Temporary + Val(MainlineOperationGUI!ConsistControlSpeed.Value) ' add the speed
        Let MainlineOperationGUI!ThreeByteD.Text = Temporary
        Let MainlineOperationGUI!FourByteD.Text = ""
        Let MainlineOperationGUI!FiveByteD.Text = ""
        Let MainlineOperationGUI!SixByteD.Text = ""
    
    End If
End If

End Sub







Private Sub PictureBoxLocmotiveCab_Click()
                   
End Sub

Private Sub PictureBoxAutomaticBrake_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxAutomaticBrake.Left = Val(PictureBoxAutomaticBrake.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxAutomaticBrake.Left = Val(PictureBoxAutomaticBrake.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxAutomaticBrake.Top = Val(PictureBoxAutomaticBrake.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxAutomaticBrake.Top = Val(PictureBoxAutomaticBrake.Top) + 1

End Sub

Private Sub PictureBoxAutomaticBrake_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxAutomaticBrake.Tag) > 0 Then
        Let PictureBoxAutomaticBrake.Tag = Trim$(Str$(Val(PictureBoxAutomaticBrake.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the automatic brake (trainline brake)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxAutomaticBrake.Tag) < 11 Then
        Let PictureBoxAutomaticBrake.Tag = Trim$(Str$(Val(PictureBoxAutomaticBrake.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the automatic brake (trainline brake)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\AutomaticBrake"
Let Temporary$ = Temporary$ + PictureBoxAutomaticBrake.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxAutomaticBrake.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxBlower_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxBlower.Left = Val(PictureBoxBlower.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxBlower.Left = Val(PictureBoxBlower.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxBlower.Top = Val(PictureBoxBlower.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxBlower.Top = Val(PictureBoxBlower.Top) + 1

End Sub

Private Sub PictureBoxBlower_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxBlower.Tag) > 0 Then
        Let PictureBoxBlower.Tag = Trim$(Str$(Val(PictureBoxBlower.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the blower."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxBlower.Tag) < 20 Then
        Let PictureBoxBlower.Tag = Trim$(Str$(Val(PictureBoxBlower.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the blower."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Blower"
Let Temporary$ = Temporary$ + PictureBoxBlower.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxBlower.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxCylinderCock_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxCylinderCock.Left = Val(PictureBoxCylinderCock.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxCylinderCock.Left = Val(PictureBoxCylinderCock.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxCylinderCock.Top = Val(PictureBoxCylinderCock.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxCylinderCock.Top = Val(PictureBoxCylinderCock.Top) + 1

End Sub

Private Sub PictureBoxCylinderCock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxCylinderCock.Tag) > 0 Then
        Let PictureBoxCylinderCock.Tag = Trim$(Str$(Val(PictureBoxCylinderCock.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the cylinder cock."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxCylinderCock.Tag) < 1 Then
        Let PictureBoxCylinderCock.Tag = Trim$(Str$(Val(PictureBoxCylinderCock.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the cylinder cock."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\CylinderCock"
Let Temporary$ = Temporary$ + PictureBoxCylinderCock.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxCylinderCock.Picture = LoadPicture(Temporary$)

End Sub

Private Sub PictureBoxDamper_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxDamper.Left = Val(PictureBoxDamper.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxDamper.Left = Val(PictureBoxDamper.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxDamper.Top = Val(PictureBoxDamper.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxDamper.Top = Val(PictureBoxDamper.Top) + 1

End Sub


Private Sub PictureBoxDamper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxDamper.Tag) > 0 Then
        Let PictureBoxDamper.Tag = Trim$(Str$(Val(PictureBoxDamper.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the damper."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxDamper.Tag) < 4 Then
        Let PictureBoxDamper.Tag = Trim$(Str$(Val(PictureBoxDamper.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the damper."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Damper"
Let Temporary$ = Temporary$ + PictureBoxDamper.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxDamper.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxFireBoxDoor_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxFireBoxDoor.Left = Val(PictureBoxFireBoxDoor.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxFireBoxDoor.Left = Val(PictureBoxFireBoxDoor.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxFireBoxDoor.Top = Val(PictureBoxFireBoxDoor.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxFireBoxDoor.Top = Val(PictureBoxFireBoxDoor.Top) + 1

End Sub


Private Sub PictureBoxFireBoxDoor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxFireBoxDoor.Tag) > 0 Then
        Let PictureBoxFireBoxDoor.Tag = Trim$(Str$(Val(PictureBoxFireBoxDoor.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the fire box door."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxFireBoxDoor.Tag) < 4 Then
        Let PictureBoxFireBoxDoor.Tag = Trim$(Str$(Val(PictureBoxFireBoxDoor.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the fire box door."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\FireBoxDoor"
Let Temporary$ = Temporary$ + PictureBoxFireBoxDoor.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxFireBoxDoor.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorSteamValveExhaust_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxInjectorSteamValveExhaust.Tag) > 0 Then
        Let PictureBoxInjectorSteamValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveExhaust.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the injector steam valve (exhaust)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxInjectorSteamValveExhaust.Tag) < 1 Then
        Let PictureBoxInjectorSteamValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveExhaust.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the injector steam valve (exhaust)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorSteamValveExhaust"
Let Temporary$ = Temporary$ + PictureBoxInjectorSteamValveExhaust.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorSteamValveExhaust.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorSteamValveLive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxInjectorSteamValveLive.Tag) > 0 Then
        Let PictureBoxInjectorSteamValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveLive.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the injector steam valve (live)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxInjectorSteamValveLive.Tag) < 1 Then
        Let PictureBoxInjectorSteamValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorSteamValveLive.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the injector steam valve (live)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorSteamValveLive"
Let Temporary$ = Temporary$ + PictureBoxInjectorSteamValveLive.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorSteamValveLive.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorWaterValveExhaust_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxInjectorWaterValveExhaust.Left = Val(PictureBoxInjectorWaterValveExhaust.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxInjectorWaterValveExhaust.Left = Val(PictureBoxInjectorWaterValveExhaust.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxInjectorWaterValveExhaust.Top = Val(PictureBoxInjectorWaterValveExhaust.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxInjectorWaterValveExhaust.Top = Val(PictureBoxInjectorWaterValveExhaust.Top) + 1

End Sub

Private Sub PictureBoxInjectorWaterValveExhaust_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If Val(PictureBoxInjectorWaterValveExhaust.Tag) > 0 Then
        Let PictureBoxInjectorWaterValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveExhaust.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path$ + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the injector water valve (exhaust)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbLeftButton Then
    If Val(PictureBoxInjectorWaterValveExhaust.Tag) < 9 Then
        Let PictureBoxInjectorWaterValveExhaust.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveExhaust.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path$ + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the injector water valve (exhaust)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorWaterValveExhaust"
Let Temporary$ = Temporary$ + PictureBoxInjectorWaterValveExhaust.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorWaterValveExhaust.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxInjectorWaterValveLive_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxInjectorWaterValveLive.Left = Val(PictureBoxInjectorWaterValveLive.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxInjectorWaterValveLive.Left = Val(PictureBoxInjectorWaterValveLive.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxInjectorWaterValveLive.Top = Val(PictureBoxInjectorWaterValveLive.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxInjectorWaterValveLive.Top = Val(PictureBoxInjectorWaterValveLive.Top) + 1

End Sub

Private Sub PictureBoxInjectorWaterValveLive_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If Val(PictureBoxInjectorWaterValveLive.Tag) > 0 Then
        Let PictureBoxInjectorWaterValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveLive.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the injector water valve (live)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbLeftButton Then
    If Val(PictureBoxInjectorWaterValveLive.Tag) < 9 Then
        Let PictureBoxInjectorWaterValveLive.Tag = Trim$(Str$(Val(PictureBoxInjectorWaterValveLive.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the injector water valve (live)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\InjectorWaterValveLive"
Let Temporary$ = Temporary$ + PictureBoxInjectorWaterValveLive.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxInjectorWaterValveLive.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxPointer_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxPointer.Left = Val(PictureBoxPointer.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxPointer.Left = Val(PictureBoxPointer.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxPointer.Top = Val(PictureBoxPointer.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxPointer.Top = Val(PictureBoxPointer.Top) + 1

End Sub


Private Sub PictureBoxRegulator_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxRegulator.Left = Val(PictureBoxRegulator.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxRegulator.Left = Val(PictureBoxRegulator.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxRegulator.Top = Val(PictureBoxRegulator.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxRegulator.Top = Val(PictureBoxRegulator.Top) + 1

End Sub

Private Sub PictureBoxRegulator_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxRegulator.Tag) > 0 Then
        Let PictureBoxRegulator.Tag = Trim$(Str$(Val(PictureBoxRegulator.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the regulator."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxRegulator.Tag) < 11 Then
        Let PictureBoxRegulator.Tag = Trim$(Str$(Val(PictureBoxRegulator.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the regulator."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Regulator"
Let Temporary$ = Temporary$ + PictureBoxRegulator.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxRegulator.Picture = LoadPicture(Temporary$)

End Sub

Private Sub PictureBoxReverser_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxReverser.Left = Val(PictureBoxReverser.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxReverser.Left = Val(PictureBoxReverser.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxReverser.Top = Val(PictureBoxReverser.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxReverser.Top = Val(PictureBoxReverser.Top) + 1

End Sub

Private Sub PictureBoxReverser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxReverser.Tag) > -15 Then
        Let PictureBoxReverser.Tag = Trim$(Str$(Val(PictureBoxReverser.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the independent brake (locomotive brake)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxReverser.Tag) < 21 Then
        Let PictureBoxReverser.Tag = Trim$(Str$(Val(PictureBoxReverser.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the independent brake (locomotive brake)."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Reverser"
Let Temporary$ = Temporary$ + PictureBoxReverser.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxReverser.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxSand_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxSand.Left = Val(PictureBoxSand.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxSand.Left = Val(PictureBoxSand.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxSand.Top = Val(PictureBoxSand.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxSand.Top = Val(PictureBoxSand.Top) + 1

End Sub

Private Sub PictureBoxSand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If Val(PictureBoxSand.Tag) > 0 Then
        Let PictureBoxSand.Tag = Trim$(Str$(Val(PictureBoxSand.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the sand lever."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbLeftButton Then
    If Val(PictureBoxSand.Tag) < 1 Then
        Let PictureBoxSand.Tag = Trim$(Str$(Val(PictureBoxSand.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the sand lever."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\Sand"
Let Temporary$ = Temporary$ + PictureBoxSand.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxSand.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxSmallInjectorCompessor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxSmallInjectorCompressor.Tag) > 0 Then
        Let PictureBoxSmallInjectorCompressor.Tag = Trim$(Str$(Val(PictureBoxSmallInjectorCompressorTag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the small injector compressor valve."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxSmallInjectorCompressor.Tag) < 1 Then
        Let PictureBoxSmallInjectorCompressor.Tag = Trim$(Str$(Val(PictureBoxSmallInjectorCompressor.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the small injector compressor valve."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\SmallInjectorCompressor"
Let Temporary$ = Temporary$ + PictureBoxSmallInjectorCompressor.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxSmallInjectorCompressor.Picture = LoadPicture(Temporary$)

End Sub


Private Sub PictureBoxSmallInjectorCompressor_KeyPress(KeyAscii As Integer)

If KeyAscii = Asc("A") Then PictureBoxSmallInjectorCompressor.Left = Val(PictureBoxSmallInjectorCompressor.Left) - 1
If KeyAscii = Asc("S") Then PictureBoxSmallInjectorCompressor.Left = Val(PictureBoxSmallInjectorCompressor.Left) + 1
If KeyAscii = Asc("W") Then PictureBoxSmallInjectorCompressor.Top = Val(PictureBoxSmallInjectorCompressor.Top) - 1
If KeyAscii = Asc("Z") Then PictureBoxSmallInjectorCompressor.Top = Val(PictureBoxSmallInjectorCompressor.Top) + 1

End Sub


Private Sub PictureBoxSmallInjectorCompressor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Val(PictureBoxSmallInjectorCompressor.Tag) > 0 Then
        Let PictureBoxSmallInjectorCompressor.Tag = Trim$(Str$(Val(PictureBoxSmallInjectorCompressor.Tag) - 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the minimum application of the small injector compressor."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
ElseIf Button = vbRightButton Then
    If Val(PictureBoxSmallInjectorCompressor.Tag) < 1 Then
        Let PictureBoxSmallInjectorCompressor.Tag = Trim$(Str$(Val(PictureBoxSmallInjectorCompressor.Tag) + 1))
        Let MainlineOperationGUI!Wave1.filename = App.Path + "\Sounds\Graphics\Control.wav"
        Let MainlineOperationGUI!Wave1.Action = wAPlay
    Else
        Let TemporaryPrompt = "You have reached the maximum application of the small injector compressor."
        MsgBox TemporaryPrompt, vbExclamation, "ATC - Engineer Error"
    End If
End If

Let Temporary$ = App.Path$
Let Temporary$ = Temporary$ + "\Graphics\Locomotive Steam1\SmallInjectorCompressor"
Let Temporary$ = Temporary$ + PictureBoxSmallInjectorCompressor.Tag
Let Temporary$ = Temporary$ + "(s1).bmp"

Let PictureBoxSmallInjectorCompressor.Picture = LoadPicture(Temporary$)

End Sub


