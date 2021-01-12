Public Enum ColorName
    Aqua = 42
    Black = 1
    Blue = 5
    BlueGray = 47
    BluePlus = 32
    BrightGreen = 4
    Brown = 53
    Coral = 22
    Cyan = 8
    DarkBlue = 11
    DarkBluePlus = 25
    DarkGreen = 51
    DarkPurple = 21
    DarkRed = 9
    DarkRedPlus = 30
    DarkTeal = 49
    DarkYellow = 12
    Gold = 44
    Gray25 = 15
    Gray40 = 48
    Gray50 = 16
    Gray80 = 56
    Green = 10
    IceBlue = 24
    Indigo = 55
    Ivory = 19
    Lavender = 39
    LightBlue = 41
    LightGreen = 35
    LightOrange = 45
    LightTurquoise = 34
    LightYellow = 36
    Lime = 43
    LiteTurquoise = 20
    OceanBlue = 23
    OliveGreen = 52
    Orange = 46
    PaleBlue = 37
    Periwinkle = 17
    Pink = 7
    PinkPlus = 26
    Plum = 54
    PlumPlus = 18
    Red = 3
    Rose = 38
    SeaGreen = 50
    SkyBlue = 33
    Tan = 40
    Teal = 14
    TealPlus = 31
    TurquoisePlus = 28
    Violet = 13
    VioletPlus = 29
    White = 2
    Yellow = 6
    YellowPlus = 27
End Enum

Public Sub DemoColorName()
	' change ActiveCell Background (Interior) Color to Aqua
	ActiveCell.Interior.ColorIndex = ColorName.Aqua
End Sub