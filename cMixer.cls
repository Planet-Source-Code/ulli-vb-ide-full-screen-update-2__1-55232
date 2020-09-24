VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Simple Sound Mixer Wrapper Class
'''''''''''''''''''''''''''''''''
'
'How To:
''''''''
'  Choose the Channel and SoundControl; this will return True if the selection was successful.
'  Get or Let the Value; note that ALL values are in %
'  For booleans (like Mute) the value 0 means False and 100 means True
'                                                       (one hundred percent true, so to say).
'
'  Success returns True when the last selection was successful.
'  ChannName returns the selected channel name (eg. Speakers).
'  CrtlName returns the selected SoundControl name (eg. VolumeControl).
'
'Authors note:
''''''''''''''
'  Unfortunately the Mixer Interface is rather complicated - maybe written by a musician (?) *g* -
'  so there are quite a few mystic API calls with plenty of params, mixer-constants with
'  ugly names, cryptic structure types, and virtual memory address pointers from one structure
'  to the next  >:-(  *grrrhh!*...
'
'  And Micro$oft's documentation is slim, to put it polite.
'
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long

Private Enum MixerConstants 'this makes them long by default

    MMSYSERR_NOERROR = 0
    MIXER_CONTROLDETAILSF_VALUE = 0
    MIXER_GETLINECONTROLSF_ONEBYTYPE = 2
    MIXER_GETLINEINFOF_COMPONENTTYPE = 3
    MIXER_SHORT_NAME_CHARS = 16
    MIXER_LONG_NAME_CHARS = 64
    MAXPNAMELEN = 32

    'mixer line constants
    MLC_DST_FIRST = 0
    MLC_SRC_FIRST = &H1000

    'Mixer control constants
    MCT_CLASS_FADER = &H50000000
    MCT_UNITS_UNSIGNED = &H30000
    MCT_FADER = MCT_CLASS_FADER Or MCT_UNITS_UNSIGNED

    MCT_CLASS_SWITCH = &H20000000
    MCT_UNITS_BOOLEAN = &H10000
    MCT_BOOLEAN = MCT_CLASS_SWITCH Or MCT_UNITS_BOOLEAN

End Enum

Public Enum Channels
    DigitalOut = MLC_DST_FIRST + 1
    LineOut = MLC_DST_FIRST + 2
    MonitorOut = MLC_DST_FIRST + 3
    SpeakersOut = MLC_DST_FIRST + 4
    HeadphonesOut = MLC_DST_FIRST + 5
    TelephoneOut = MLC_DST_FIRST + 6
    WaveInOut = MLC_DST_FIRST + 7
    VoiceInOut = MLC_DST_FIRST + 8
    DigitalIn = MLC_SRC_FIRST + 1
    LineIn = MLC_SRC_FIRST + 2
    MikrophoneIn = MLC_SRC_FIRST + 3
    SynthesizerIn = MLC_SRC_FIRST + 4
    CompactDiscIn = MLC_SRC_FIRST + 5
    TelephoneIn = MLC_SRC_FIRST + 6
    PCSpeakerIn = MLC_SRC_FIRST + 7
    WaveOutIn = MLC_SRC_FIRST + 8
    AuxiliaryIn = MLC_SRC_FIRST + 9
    AnalogIn = MLC_SRC_FIRST + 10
End Enum
#If False Then
Private DigitalOut, LineOut, MonitorOut, SpeakersOut, HeadphonesOut, TelephoneOut, WaveInOut, VoiceInOut
Private DigitalIn, LineIn, MikrophoneIn, SynthesizerIn, CompactDiscIn, TelephoneIn, PCSpeakerIn, WaveOutIn, AuxiliaryIn, AnalogIn
#End If

Public Enum SoundControls
    Mute = MCT_BOOLEAN + 2
    Mono = MCT_BOOLEAN + 3
    Loudness = MCT_BOOLEAN + 4
    StereoEnhance = MCT_BOOLEAN + 5
    Volume = MCT_FADER + 1
    Bass = MCT_FADER + 2
    Treble = MCT_FADER + 3
    Equalizer = MCT_FADER + 4
End Enum
#If False Then
Private Loudness, Mute, StereoEnhance, Mono, Pan, Volume, Bass, Treble, Equalizer
#End If

'mixer handle
Private hMixer As Long

'mixer structures
Private Type MIXERLINE
    cbStruct            As Long 'size in bytes of MIXERLINE structure
    dwDestination       As Long 'zero based destination index
    dwSource            As Long 'zero based source index (if source)
    dwLineID            As Long 'unique line id for mixer device
    fdwLine             As Long 'state/information about line
    dwUser              As Long 'driver specific information
    dwComponentType     As Long 'component type line connects to
    cChannels           As Long 'number of channels line supports
    cConnections        As Long 'number of connections (possible)
    cControls           As Long 'number of controls at this line
    szShortName(1 To MIXER_SHORT_NAME_CHARS)    As Byte
    szName(1 To MIXER_LONG_NAME_CHARS)          As Byte
    dwType              As Long
    dwDeviceID          As Long
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname(1 To MAXPNAMELEN)  As Byte
End Type
Private ChannelLine As MIXERLINE

Private Type MIXERLINECONTROLS
    cbStruct            As Long 'size in Byte of MIXERLINECONTROLS
    dwLineID            As Long 'line id (from MIXERLINE.dwLineID)
    dwControl           As Long 'MIXER_GETLINECONTROLSF_ONEBYID or MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls           As Long 'count of controls pamxctrl points to
    cbmxctrl            As Long 'size in Byte of _one_ MIXERCONTROL
    pamxctrl            As Long 'pointer to first MIXERCONTROL array
End Type
Private ChannelControls As MIXERLINECONTROLS

Private Type MIXERCONTROL
    cbStruct            As Long 'size in Byte of MIXERCONTROL
    dwControlID         As Long 'unique control id for mixer device
    dwControlType       As Long 'MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl          As Long 'MIXERCONTROL_CONTROLF_xxx
    cMultipleItems      As Long 'if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName(1 To MIXER_SHORT_NAME_CHARS)   As Byte 'short name of control
    szName(1 To MIXER_LONG_NAME_CHARS)         As Byte 'long name of control
    lMinimum            As Long 'Minimum value
    lMaximum            As Long 'Maximum value
    reserved(10)        As Long 'reserved structure space
End Type
Private ValueControl As MIXERCONTROL

Private Type MIXERCONTROLDETAILS
    cbStruct            As Long 'size in Byte of MIXERCONTROLDETAILS
    dwControlID         As Long 'control id to get/set details on
    cChannels           As Long 'number of channels in paDetails array
    item                As Long 'hwndOwner or cMultipleItems
    cbDetails           As Long 'size of one details_XX struct
    paDetails           As Long 'pointer to array of details_XX structs
End Type
Private ControlDetails As MIXERCONTROLDETAILS

'Properties
Private myValue         As Long
Private myMinValue      As Long
Private myMaxValue      As Long
Private mySuccess       As Boolean
Private myChannName     As String
Private myCtrlName      As String

Public Property Get ChannelName() As String

  'returns the chosen channel name

    ChannelName = Left$(myChannName, InStr(myChannName, Chr$(0)) - 1)

End Property

Public Function Choose(Channel As Channels, SoundControl As SoundControls) As Boolean

    mySuccess = CBool(hMixer)
    If mySuccess Then
        myChannName = Chr$(0)
        myCtrlName = myChannName
        With ChannelLine
            .cbStruct = Len(ChannelLine)
            .dwComponentType = Channel
        End With 'CHANNELLINE
        If mixerGetLineInfo(hMixer, ChannelLine, MIXER_GETLINEINFOF_COMPONENTTYPE) = MMSYSERR_NOERROR Then
            myCtrlName = StrConv(ChannelLine.szName, vbUnicode)
            With ChannelControls
                .cbStruct = Len(ChannelControls)
                .dwLineID = ChannelLine.dwLineID
                .dwControl = SoundControl
                .cControls = 1
                .cbmxctrl = Len(ValueControl)
                .pamxctrl = VarPtr(ValueControl)
            End With 'CHANNELCONTROLS
            If mixerGetLineControls(hMixer, ChannelControls, MIXER_GETLINECONTROLSF_ONEBYTYPE) = MMSYSERR_NOERROR Then
                With ValueControl
                    .cbStruct = Len(ValueControl)
                    myMinValue = .lMinimum
                    myMaxValue = .lMaximum
                    myChannName = StrConv(.szName, vbUnicode)
                End With 'VALUECONTROL
              Else 'NOT MIXERGETLINECONTROLS(HMIXER,...
                mySuccess = False
            End If
          Else 'NOT MIXERGETLINEINFO(HMIXER,...
            mySuccess = False
        End If
    End If
    Choose = mySuccess

End Function

Private Sub Class_Initialize()

    TidyUp
    mixerOpen hMixer, 0&, 0&, 0&, 0&
    Choose SpeakersOut, Volume 'preset path to speaker volume

End Sub

Private Sub Class_Terminate()

    mixerClose hMixer
    TidyUp

End Sub

Public Property Get ControlName() As String

  'returns the chosen sound control name

    ControlName = Left$(myCtrlName, InStr(myCtrlName, Chr$(0)) - 1)

End Property

Private Sub SetUpControlDetails(ByRef Value As Long)

    With ControlDetails
        .cbStruct = Len(ControlDetails)
        .item = 0
        .dwControlID = ValueControl.dwControlID
        .cChannels = 1
        .cbDetails = Len(Value)
        .paDetails = VarPtr(Value)
    End With 'ControlDetails

End Sub

Public Property Get Success() As Boolean

  'returns success of last choice

    Success = mySuccess

End Property

Private Sub TidyUp()

    myValue = 0
    myMinValue = 0
    myMaxValue = 0
    mySuccess = False
    myChannName = vbNullString
    myCtrlName = vbNullString
    hMixer = 0

End Sub

Public Property Get Value() As Long

    If hMixer Then
        SetUpControlDetails myValue
        mixerGetControlDetails hMixer, ControlDetails, MIXER_CONTROLDETAILSF_VALUE
        On Error Resume Next 'in case myMaxValue and myMinValue are equal, causing a divide by zero
            Value = (myValue - myMinValue) * 100 / (myMaxValue - myMinValue) 'convert to %
            If Err Then
                Value = 0
            End If
        On Error GoTo 0
    End If

End Property

Public Property Let Value(ByVal Percent As Long)

    If hMixer Then
        If Percent >= 0 And Percent <= 100 Then
            myValue = (myMaxValue - myMinValue) * Percent / 100 + myMinValue 'convert to %
            SetUpControlDetails myValue
            mixerSetControlDetails hMixer, ControlDetails, MIXER_CONTROLDETAILSF_VALUE
        End If
    End If

End Property

':) Ulli's VB Code Formatter V2.17.4 (2004-Aug-02 11:20) 163 + 131 = 294 Lines
