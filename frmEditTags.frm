VERSION 5.00
Begin VB.Form frmEditTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Tags"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmEditTags.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbGenre 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtEditTags 
      Height          =   285
      Index           =   4
      Left            =   240
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtEditTags 
      Height          =   285
      Index           =   3
      Left            =   240
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtEditTags 
      Height          =   285
      Index           =   2
      Left            =   240
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtEditTags 
      Height          =   285
      Index           =   1
      Left            =   240
      MaxLength       =   30
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtEditTags 
      Height          =   285
      Index           =   0
      Left            =   240
      MaxLength       =   30
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblEditTags 
      Caption         =   "Genre:"
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblEditTags 
      Caption         =   "Comment:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblEditTags 
      Caption         =   "Year:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblEditTags 
      Caption         =   "Album:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblEditTags 
      Caption         =   "Artist:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblEditTags 
      Caption         =   "Tag:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmEditTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mFullPath As String
Dim mTag As String
Dim mArtist As String
Dim mAlbum As String
Dim mYear As String
Dim mComment As String
Dim mGenre As Integer
Dim b() As Byte
Dim TagPresent As Boolean
Dim Length As Long

Dim CommentStart As Long
Dim CommentEnd As Long
Dim YearStart As Long
Dim YearEnd As Long
Dim AlbumStart As Long
Dim AlbumEnd As Long
Dim ArtistStart As Long
Dim ArtistEnd As Long
Dim TagStart As Long
Dim TagEnd As Long
Dim GenreStartEnd As Long
Dim TagIndicatorStart As Long

Private GenreArray() As String         ' we use this array to fill all the Genre's ( look in form load)

Private Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"

Private Sub cmdCancel_Click()
    On Error Resume Next
    
    'Unload the form
    Unload Me

End Sub

Private Sub cmdOK_Click()
    'On Error GoTo errHandler
    
    Dim i As Long
    Dim m As New clsMousePointer
    m.SetCursor
    
    'Pad with spaces
    mTag = txtEditTags(0).Text & Space(30 - Len(txtEditTags(0).Text))
    mArtist = txtEditTags(1).Text & Space(30 - Len(txtEditTags(1).Text))
    mAlbum = txtEditTags(2).Text & Space(30 - Len(txtEditTags(2).Text))
    mYear = txtEditTags(3).Text & Space(4 - Len(txtEditTags(3).Text))
    mComment = txtEditTags(4).Text & Space(30 - Len(txtEditTags(4).Text))
    mGenre = cmbGenre.ListIndex
    
    If Not TagPresent Then
        'Create a tag
        
        Length = Length + 127
        ReDim Preserve b(Length)
        
        'Set the boundaries
        CommentStart = Length - 31
        CommentEnd = Length - 2
        YearStart = Length - 35
        YearEnd = Length - 32
        AlbumStart = Length - 65
        AlbumEnd = Length - 36
        ArtistStart = Length - 95
        ArtistEnd = Length - 66
        TagStart = Length - 125
        TagEnd = Length - 96
        TagIndicatorStart = Length - 128
        GenreStartEnd = Length - 1
        
        'Set the tag indicator
        b(TagIndicatorStart) = Asc("T")
        b(TagIndicatorStart + 1) = Asc("A")
        b(TagIndicatorStart + 2) = Asc("G")
        
    End If
    
    'Set the comment
    For i = CommentStart To CommentEnd
        b(i) = Asc(Mid(mComment, 30 - (CommentEnd - i), 1))
    Next i

    'Set the year
    For i = YearStart To YearEnd
        b(i) = Asc(Mid(mYear, 4 - (YearEnd - i), 1))
    Next i

    'Set the album
    For i = AlbumStart To AlbumEnd
        b(i) = Asc(Mid(mAlbum, 30 - (AlbumEnd - i), 1))
    Next i

    'Set the artist
    For i = ArtistStart To ArtistEnd
        b(i) = Asc(Mid(mArtist, 30 - (ArtistEnd - i), 1))
    Next i
    
    'Set the tag
    For i = TagStart To TagEnd
        b(i) = Asc(Mid(mTag, 30 - (TagEnd - i), 1))
    Next i
    
    'Set the genre
    b(GenreStartEnd) = CByte(mGenre)
    
    'This will write the encrypted\decrypted file
    ReDim Preserve b(Length - 1)
    Open mFullPath For Binary Access Write As #1
        Put #1, , b
    Close #1
    
    Unload Me
    
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical

    Unload Me
    
End Sub

Public Sub StartEdit(FullPath As String)

    Dim i As Integer

    mFullPath = FullPath
    
    GenreArray = Split(sGenreMatrix, "|")   ' we fill the array with the Genre's
    For i = LBound(GenreArray) To UBound(GenreArray)
        cmbGenre.AddItem GenreArray(i)        ' now fill the Combobox with the array, and voila, the code you
    Next i
    
    If GetTags(mFullPath) Then
    
        txtEditTags(0).Text = mTag
        txtEditTags(1).Text = mArtist
        txtEditTags(2).Text = mAlbum
        txtEditTags(3).Text = mYear
        txtEditTags(4).Text = mComment
        cmbGenre.ListIndex = mGenre
        
        Me.Show 1
    End If

End Sub

Public Function GetTags(FullPath As String) As Boolean
    'On Error GoTo errHandler
    
    Dim i As Long
    Dim Tag As String
    Dim Artist As String
    Dim Album As String
    Dim Year As String
    Dim Comment As String
    Dim TagIndicator As String
    Dim m As New clsMousePointer
    m.SetCursor
    
    Length = FileLen(FullPath)
    ReDim b(0)
    
    If Length = 0 Then
        Err.Raise 10000, , FullPath & " does not contain any data."
    End If
    
    'Set the boundaries
    CommentStart = Length - 31
    CommentEnd = Length - 2
    YearStart = Length - 35
    YearEnd = Length - 32
    AlbumStart = Length - 65
    AlbumEnd = Length - 36
    ArtistStart = Length - 95
    ArtistEnd = Length - 66
    TagStart = Length - 125
    TagEnd = Length - 96
    TagIndicatorStart = Length - 128
    GenreStartEnd = Length - 1
    
    ReDim b(Length)

    Open FullPath For Binary Access Read As #1
        Get #1, , b()
    Close #1

    'See if there is a tag
    TagIndicator = Chr(b(TagIndicatorStart))
    TagIndicator = TagIndicator & Chr(b(TagIndicatorStart + 1))
    TagIndicator = TagIndicator & Chr(b(TagIndicatorStart + 2))
    

    If TagIndicator <> "TAG" Then
        'No tag Present
        mTag = frmMP3.GetSong(FullPath)
        mArtist = ""
        mAlbum = ""
        mYear = ""
        mComment = ""
        mGenre = -1
        TagPresent = False
        GetTags = True
        Exit Function
    Else
        TagPresent = True
    End If
    
    'Get the comment
    For i = CommentStart To CommentEnd
        Comment = Comment & Chr(b(i))
    Next i

    mComment = Trim(Comment)

    'Get the Year
    For i = YearStart To YearEnd
        Year = Year & Chr(b(i))
    Next i

    mYear = Trim(Year)

    'Get the Album
    For i = AlbumStart To AlbumEnd
        Album = Album & Chr(b(i))
    Next i

    mAlbum = Trim(Album)

    'Get the Artist
    For i = ArtistStart To ArtistEnd
        Artist = Artist & Chr(b(i))
    Next i

    mArtist = Trim(Artist)
    
    'Get the Tag
    For i = TagStart To TagEnd
        Tag = Tag & Chr(b(i))
    Next i

    mTag = Trim(Tag)
    
    'Get the Genre
    mGenre = b(GenreStartEnd)
    
    GetTags = True
    
    Exit Function
    
errHandler:
    MsgBox Err.Description, vbCritical
    
    GetTags = False
    
End Function

Private Sub txtEditTags_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 3 Then
        If KeyAscii <> 8 Then
            If Not IsNumeric(Chr(KeyAscii)) Then
                 KeyAscii = 0
            End If
        End If
    End If

End Sub

Private Sub txtEditTags_LostFocus(Index As Integer)

    txtEditTags(Index).Text = Trim(txtEditTags(Index).Text)

End Sub
