using System.Text;

namespace DocxportNet.symbols;

public static class WebdingsEncoding
{
	public static readonly string?[] Table = [
	    "\u0000", // 0x00
        "\u0001", // 0x01
        "\u0002", // 0x02
        "\u0003", // 0x03
        "\u0004", // 0x04
        "\u0005", // 0x05
        "\u0006", // 0x06
        "\u0007", // 0x07
        "\u0008", // 0x08
        "\u0009", // 0x09
        "\u000A", // 0x0A
        "\u000B", // 0x0B
        "\u000C", // 0x0C
        "\u000D", // 0x0D
        "\u000E", // 0x0E
        "\u000F", // 0x0F
        "\u0010", // 0x10
        "\u0011", // 0x11
        "\u0012", // 0x12
        "\u0013", // 0x13
        "\u0014", // 0x14
        "\u0015", // 0x15
        "\u0016", // 0x16
        "\u0017", // 0x17
        "\u0018", // 0x18
        "\u0019", // 0x19
        "\u001A", // 0x1A
        "\u001B", // 0x1B
        "\u001C", // 0x1C
        "\u001D", // 0x1D
        "\u001E", // 0x1E
        "\u001F", // 0x1F
        " ", // 0x20
        "🕷", // 0x21
        "🕸", // 0x22
        "🕲", // 0x23
        "🕶", // 0x24
        "🏆", // 0x25
        "🎖", // 0x26
        "🖇", // 0x27
        "🗨", // 0x28
        "🗩", // 0x29
        "🗰", // 0x2A
        "🗱", // 0x2B
        "🌶", // 0x2C
        "🎗", // 0x2D
        "🙾", // 0x2E
        "🙼", // 0x2F
        "🗕", // 0x30
        "🗖", // 0x31
        "🗗", // 0x32
        "⏴", // 0x33
        "⏵", // 0x34
        "⏶", // 0x35
        "⏷", // 0x36
        "⏪", // 0x37
        "⏩", // 0x38
        "⏮", // 0x39
        "⏭", // 0x3A
        "⏸", // 0x3B
        "⏹", // 0x3C
        "⏺", // 0x3D
        "🗚", // 0x3E
        "🗳", // 0x3F
        "🛠", // 0x40
        "🏗", // 0x41
        "🏘", // 0x42
        "🏙", // 0x43
        "🏚", // 0x44
        "🏜", // 0x45
        "🏭", // 0x46
        "🏛", // 0x47
        "🏠", // 0x48
        "🏖", // 0x49
        "🏝", // 0x4A
        "🛣", // 0x4B
        "🔍", // 0x4C
        "🏔", // 0x4D
        "👁", // 0x4E
        "👂", // 0x4F
        "🏞", // 0x50
        "🏕", // 0x51
        "🛤", // 0x52
        "🏟", // 0x53
        "🛳", // 0x54
        "🕬", // 0x55
        "🕫", // 0x56
        "🕨", // 0x57
        "🔈", // 0x58
        "🎔", // 0x59
        "🎕", // 0x5A
        "🗬", // 0x5B
        "🙽", // 0x5C
        "🗭", // 0x5D
        "🗪", // 0x5E
        "🗫", // 0x5F
        "⮔", // 0x60
        "✔", // 0x61
        "🚲", // 0x62
        "□", // 0x63
        "🛡", // 0x64
        "📦", // 0x65
        "🛱", // 0x66
        "■", // 0x67
        "🚑", // 0x68
        "🛈", // 0x69
        "🛩", // 0x6A
        "🛰", // 0x6B
        "🟈", // 0x6C
        "🕴", // 0x6D
        "⚫", // 0x6E
        "🛥", // 0x6F
        "🚔", // 0x70
        "🗘", // 0x71
        "🗙", // 0x72
        "❓", // 0x73
        "🛲", // 0x74
        "🚇", // 0x75
        "🚍", // 0x76
        "⛳", // 0x77
        "🛇", // 0x78
        "⊖", // 0x79
        "🚭", // 0x7A
        "🗮", // 0x7B
        "|", // 0x7C
        "🗯", // 0x7D
        "🗲", // 0x7E
        "\u007F", // 0x7F
        "🚹", // 0x80
        "🚺", // 0x81
        "🛉", // 0x82
        "🛊", // 0x83
        "🚼", // 0x84
        "👽", // 0x85
        "🏋", // 0x86
        "⛷", // 0x87
        "🏂", // 0x88
        "🏌", // 0x89
        "🏊", // 0x8A
        "🏄", // 0x8B
        "🏍", // 0x8C
        "🏎", // 0x8D
        "🚘", // 0x8E
        "🗠", // 0x8F
        "🛢", // 0x90
        "💰", // 0x91
        "🏷", // 0x92
        "💳", // 0x93
        "👪", // 0x94
        "🗡", // 0x95
        "🗢", // 0x96
        "🗣", // 0x97
        "✯", // 0x98
        "🖄", // 0x99
        "🖅", // 0x9A
        "🖃", // 0x9B
        "🖆", // 0x9C
        "🖹", // 0x9D
        "🖺", // 0x9E
        "🖻", // 0x9F
        "🕵", // 0xA0
        "🕰", // 0xA1
        "🖽", // 0xA2
        "🖾", // 0xA3
        "📋", // 0xA4
        "🗒", // 0xA5
        "🗓", // 0xA6
        "📖", // 0xA7
        "📚", // 0xA8
        "🗞", // 0xA9
        "🗟", // 0xAA
        "🗃", // 0xAB
        "🗂", // 0xAC
        "🖼", // 0xAD
        "🎭", // 0xAE
        "🎜", // 0xAF
        "🎘", // 0xB0
        "🎙", // 0xB1
        "🎧", // 0xB2
        "💿", // 0xB3
        "🎞", // 0xB4
        "📷", // 0xB5
        "🎟", // 0xB6
        "🎬", // 0xB7
        "📽", // 0xB8
        "📹", // 0xB9
        "📾", // 0xBA
        "📻", // 0xBB
        "🎚", // 0xBC
        "🎛", // 0xBD
        "📺", // 0xBE
        "💻", // 0xBF
        "🖥", // 0xC0
        "🖦", // 0xC1
        "🖧", // 0xC2
        "🕹", // 0xC3
        "🎮", // 0xC4
        "🕻", // 0xC5
        "🕼", // 0xC6
        "📟", // 0xC7
        "🖁", // 0xC8
        "🖀", // 0xC9
        "🖨", // 0xCA
        "🖩", // 0xCB
        "🖿", // 0xCC
        "🖪", // 0xCD
        "🗜", // 0xCE
        "🔒", // 0xCF
        "🔓", // 0xD0
        "🗝", // 0xD1
        "📥", // 0xD2
        "📤", // 0xD3
        "🕳", // 0xD4
        "🌣", // 0xD5
        "🌤", // 0xD6
        "🌥", // 0xD7
        "🌦", // 0xD8
        "☁", // 0xD9
        "🌧", // 0xDA
        "🌨", // 0xDB
        "🌩", // 0xDC
        "🌪", // 0xDD
        "🌬", // 0xDE
        "🌫", // 0xDF
        "🌜", // 0xE0
        "🌡", // 0xE1
        "🛋", // 0xE2
        "🛏", // 0xE3
        "🍽", // 0xE4
        "🍸", // 0xE5
        "🛎", // 0xE6
        "🛍", // 0xE7
        "Ⓟ", // 0xE8
        "♿", // 0xE9
        "🛆", // 0xEA
        "🖈", // 0xEB
        "🎓", // 0xEC
        "🗤", // 0xED
        "🗥", // 0xEE
        "🗦", // 0xEF
        "🗧", // 0xF0
        "🛪", // 0xF1
        "🐿", // 0xF2
        "🐦", // 0xF3
        "🐟", // 0xF4
        "🐕", // 0xF5
        "🐈", // 0xF6
        "🙬", // 0xF7
        "🙮", // 0xF8
        "🙭", // 0xF9
        "🙯", // 0xFA
        "🗺", // 0xFB
        "🌍", // 0xFC
        "🌏", // 0xFD
        "🌎", // 0xFE
        "🕊" // 0xFF
    ];

    // Key = Webdings code (byte 0x00..0xFF as used by the Webdings font)
    // Value = Unicode string (encode with UTF-8 as needed)
    public static readonly Dictionary<byte, string> WebdingsToUnicode = new() {
		[0x20] = "\u0020", // ‘ ’ U+0020 Space
		[0x21] = "\U0001F577", // 🕷 U+1F577 Spider
		[0x22] = "\U0001F578", // 🕸 U+1F578 Spider web
		[0x23] = "\U0001F572", // 🕲 U+1F572 No piracy
		[0x24] = "\U0001F576", // 🕶 U+1F576 Dark sunglasses
		[0x25] = "\U0001F3C6", // 🏆 U+1F3C6 Trophy
		[0x26] = "\U0001F396", // 🎖 U+1F396 Military medal
		[0x27] = "\U0001F587", // 🖇 U+1F587 Linked paperclips
		[0x28] = "\U0001F5E8", // 🗨 U+1F5E8 Left speech bubble
		[0x29] = "\U0001F5E9", // 🗩 U+1F5E9 Right speech bubble
		[0x2A] = "\U0001F5F0", // 🗰 U+1F5F0 Mood bubble
		[0x2B] = "\U0001F5F1", // 🗱 U+1F5F1 Lightning mood bubble
		[0x2C] = "\U0001F336", // 🌶 U+1F336 Hot pepper
		[0x2D] = "\U0001F397", // 🎗 U+1F397 Reminder ribbon
		[0x2E] = "\U0001F67E", // 🙾 U+1F67E Checker board
		[0x2F] = "\U0001F67C", // 🙼 U+1F67C Very heavy solidus
		[0x30] = "\U0001F5D5", // 🗕 U+1F5D5 Minimize
		[0x31] = "\U0001F5D6", // 🗖 U+1F5D6 Maximize
		[0x32] = "\U0001F5D7", // 🗗 U+1F5D7 Overlap
		[0x33] = "\u23F4", // ⏴ U+23F4 Black medium left-pointing triangle
		[0x34] = "\u23F5", // ⏵ U+23F5 Black medium right-pointing triangle
		[0x35] = "\u23F6", // ⏶ U+23F6 Black medium up-pointing triangle
		[0x36] = "\u23F7", // ⏷ U+23F7 Black medium down-pointing triangle
		[0x37] = "\u23EA", // ⏪ U+23EA Black left-pointing double triangle
		[0x38] = "\u23E9", // ⏩ U+23E9 Black right-pointing double triangle
		[0x39] = "\u23EE", // ⏮ U+23EE Black left-pointing double triangle with vertical bar
		[0x3A] = "\u23ED", // ⏭ U+23ED Black right-pointing double triangle with vertical bar
		[0x3B] = "\u23F8", // ⏸ U+23F8 Double vertical bar
		[0x3C] = "\u23F9", // ⏹ U+23F9 Black square for stop
		[0x3D] = "\u23FA", // ⏺ U+23FA Black circle for record
		[0x3E] = "\U0001F5DA", // 🗚 U+1F5DA Increase font size symbol
		[0x3F] = "\U0001F5F3", // 🗳 U+1F5F3 Ballot box with ballot
		[0x40] = "\U0001F6E0", // 🛠 U+1F6E0 Hammer and wrench
		[0x41] = "\U0001F3D7", // 🏗 U+1F3D7 Building construction
		[0x42] = "\U0001F3D8", // 🏘 U+1F3D8 House buildings
		[0x43] = "\U0001F3D9", // 🏙 U+1F3D9 Cityscape
		[0x44] = "\U0001F3DA", // 🏚 U+1F3DA Derelict house building
		[0x45] = "\U0001F3DC", // 🏜 U+1F3DC Desert
		[0x46] = "\U0001F3ED", // 🏭 U+1F3ED Factory
		[0x47] = "\U0001F3DB", // 🏛 U+1F3DB Classical building
		[0x48] = "\U0001F3E0", // 🏠 U+1F3E0 House building
		[0x49] = "\U0001F3D6", // 🏖 U+1F3D6 Beach with umbrella
		[0x4A] = "\U0001F3DD", // 🏝 U+1F3DD Desert island
		[0x4B] = "\U0001F6E3", // 🛣 U+1F6E3 Motorway
		[0x4C] = "\U0001F50D", // 🔍 U+1F50D Left-pointing magnifying glass
		[0x4D] = "\U0001F3D4", // 🏔 U+1F3D4 Snow capped mountain
		[0x4E] = "\U0001F441", // 👁 U+1F441 Eye
		[0x4F] = "\U0001F442", // 👂 U+1F442 Ear
		[0x50] = "\U0001F3DE", // 🏞 U+1F3DE National park
		[0x51] = "\U0001F3D5", // 🏕 U+1F3D5 Camping
		[0x52] = "\U0001F6E4", // 🛤 U+1F6E4 Railway track
		[0x53] = "\U0001F3DF", // 🏟 U+1F3DF Stadium
		[0x54] = "\U0001F6F3", // 🛳 U+1F6F3 Passenger ship
		[0x55] = "\U0001F56C", // 🕬 U+1F56C Bullhorn with sound waves
		[0x56] = "\U0001F56B", // 🕫 U+1F56B Bullhorn
		[0x57] = "\U0001F568", // 🕨 U+1F568 Right speaker
		[0x58] = "\U0001F508", // 🔈 U+1F508 Speaker
		[0x59] = "\U0001F394", // 🎔 U+1F394 Heart with tip on the left
		[0x5A] = "\U0001F395", // 🎕 U+1F395 Bouquet of flowers
		[0x5B] = "\U0001F5EC", // 🗬 U+1F5EC Left thought bubble
		[0x5C] = "\U0001F67D", // 🙽 U+1F67D Very heavy reverse solidus
		[0x5D] = "\U0001F5ED", // 🗭 U+1F5ED Right thought bubble
		[0x5E] = "\U0001F5EA", // 🗪 U+1F5EA Two speech bubbles
		[0x5F] = "\U0001F5EB", // 🗫 U+1F5EB Three speech bubbles
		[0x60] = "\u2B94", // ⮔ U+2B94 Four corner arrows circling anticlockwise
		[0x61] = "\u2714", // ✔ U+2714 Heavy check mark
		[0x62] = "\U0001F6B2", // 🚲 U+1F6B2 Bicycle
		[0x63] = "\u25A1", // □ U+25A1 White square
		[0x64] = "\U0001F6E1", // 🛡 U+1F6E1 Shield
		[0x65] = "\U0001F4E6", // 📦 U+1F4E6 Package
		[0x66] = "\U0001F6F1", // 🛱 U+1F6F1 Oncoming fire engine
		[0x67] = "\u25A0", // ■ U+25A0 Black square
		[0x68] = "\U0001F691", // 🚑 U+1F691 Ambulance
		[0x69] = "\U0001F6C8", // 🛈 U+1F6C8 Circled information source
		[0x6A] = "\U0001F6E9", // 🛩 U+1F6E9 Small airplane
		[0x6B] = "\U0001F6F0", // 🛰 U+1F6F0 Satellite
		[0x6C] = "\U0001F7C8", // 🟈 U+1F7C8 Reverse light four pointed pinwheel star
		[0x6D] = "\U0001F574", // 🕴 U+1F574 Man in business suit levitating
		[0x6E] = "\u26AB", // ⚫ U+26AB Medium black circle
		[0x6F] = "\U0001F6E5", // 🛥 U+1F6E5 Motor boat
		[0x70] = "\U0001F694", // 🚔 U+1F694 Oncoming police car
		[0x71] = "\U0001F5D8", // 🗘 U+1F5D8 Clockwise right and left semicircle arrows
		[0x72] = "\U0001F5D9", // 🗙 U+1F5D9 Cancellation X
		[0x73] = "\u2753", // ❓ U+2753 Black question mark ornament
		[0x74] = "\U0001F6F2", // 🛲 U+1F6F2 Diesel locomotive
		[0x75] = "\U0001F687", // 🚇 U+1F687 Metro
		[0x76] = "\U0001F68D", // 🚍 U+1F68D Oncoming bus
		[0x77] = "\u26F3", // ⛳ U+26F3 Flag in hole
		[0x78] = "\U0001F6C7", // 🛇 U+1F6C7 Prohibited sign
		[0x79] = "\u2296", // ⊖ U+2296 Circled minus
		[0x7A] = "\U0001F6AD", // 🚭 U+1F6AD No smoking symbol
		[0x7B] = "\U0001F5EE", // 🗮 U+1F5EE Left anger bubble
		[0x7C] = "\u007C", // | U+007C Vertical line
		[0x7D] = "\U0001F5EF", // 🗯 U+1F5EF Right anger bubble
		[0x7E] = "\U0001F5F2", // 🗲 U+1F5F2 Lightning mood
							   // 0x7F: no mapping in the Webdings table
		[0x80] = "\U0001F6B9", // 🚹 U+1F6B9 Mens symbol
		[0x81] = "\U0001F6BA", // 🚺 U+1F6BA Womens symbol
		[0x82] = "\U0001F6C9", // 🛉 U+1F6C9 Boys symbol
		[0x83] = "\U0001F6CA", // 🛊 U+1F6CA Girls symbol
		[0x84] = "\U0001F6BC", // 🚼 U+1F6BC Baby symbol
		[0x85] = "\U0001F47D", // 👽 U+1F47D Extraterrestrial alien
		[0x86] = "\U0001F3CB", // 🏋 U+1F3CB Weight lifter
		[0x87] = "\u26F7", // ⛷ U+26F7 Skier
		[0x88] = "\U0001F3C2", // 🏂 U+1F3C2 Snowboarder
		[0x89] = "\U0001F3CC", // 🏌 U+1F3CC Golfer
		[0x8A] = "\U0001F3CA", // 🏊 U+1F3CA Swimmer
		[0x8B] = "\U0001F3C4", // 🏄 U+1F3C4 Surfer
		[0x8C] = "\U0001F3CD", // 🏍 U+1F3CD Racing motorcycle
		[0x8D] = "\U0001F3CE", // 🏎 U+1F3CE Racing car
		[0x8E] = "\U0001F698", // 🚘 U+1F698 Oncoming automobile
		[0x8F] = "\U0001F5E0", // 🗠 U+1F5E0 Stock chart
		[0x90] = "\U0001F6E2", // 🛢 U+1F6E2 Oil drum
		[0x91] = "\U0001F4B0", // 💰 U+1F4B0 Money bag
		[0x92] = "\U0001F3F7", // 🏷 U+1F3F7 Label
		[0x93] = "\U0001F4B3", // 💳 U+1F4B3 Credit card
		[0x94] = "\U0001F46A", // 👪 U+1F46A Family
		[0x95] = "\U0001F5E1", // 🗡 U+1F5E1 Dagger knife
		[0x96] = "\U0001F5E2", // 🗢 U+1F5E2 Lips
		[0x97] = "\U0001F5E3", // 🗣 U+1F5E3 Speaking head in silhouette
		[0x98] = "\u272F", // ✯ U+272F Pinwheel star
		[0x99] = "\U0001F584", // 🖄 U+1F584 Envelope with lightning
		[0x9A] = "\U0001F585", // 🖅 U+1F585 Flying envelope
		[0x9B] = "\U0001F583", // 🖃 U+1F583 Stamped envelope
		[0x9C] = "\U0001F586", // 🖆 U+1F586 Pen over stamped envelope
		[0x9D] = "\U0001F5B9", // 🖹 U+1F5B9 Document with text
		[0x9E] = "\U0001F5BA", // 🖺 U+1F5BA Document with text and picture
		[0x9F] = "\U0001F5BB", // 🖻 U+1F5BB Document with picture
		[0xA0] = "\U0001F575", // 🕵 U+1F575 Sleuth or spy
		[0xA1] = "\U0001F570", // 🕰 U+1F570 Mantelpiece clock
		[0xA2] = "\U0001F5BD", // 🖽 U+1F5BD Frame with tiles
		[0xA3] = "\U0001F5BE", // 🖾 U+1F5BE Frame with an X
		[0xA4] = "\U0001F4CB", // 📋 U+1F4CB Clipboard
		[0xA5] = "\U0001F5D2", // 🗒 U+1F5D2 Spiral note pad
		[0xA6] = "\U0001F5D3", // 🗓 U+1F5D3 Spiral calendar pad
		[0xA7] = "\U0001F4D6", // 📖 U+1F4D6 Open book
		[0xA8] = "\U0001F4DA", // 📚 U+1F4DA Books
		[0xA9] = "\U0001F5DE", // 🗞 U+1F5DE Rolled-up newspaper
		[0xAA] = "\U0001F5DF", // 🗟 U+1F5DF Page with circled text
		[0xAB] = "\U0001F5C3", // 🗃 U+1F5C3 Card file box
		[0xAC] = "\U0001F5C2", // 🗂 U+1F5C2 Card index dividers
		[0xAD] = "\U0001F5BC", // 🖼 U+1F5BC Frame with picture
		[0xAE] = "\U0001F3AD", // 🎭 U+1F3AD Performing arts
		[0xAF] = "\U0001F39C", // 🎜 U+1F39C Beamed ascending musical notes
		[0xB0] = "\U0001F398", // 🎘 U+1F398 Musical keyboard with jacks
		[0xB1] = "\U0001F399", // 🎙 U+1F399 Studio microphone
		[0xB2] = "\U0001F3A7", // 🎧 U+1F3A7 Headphone
		[0xB3] = "\U0001F4BF", // 💿 U+1F4BF Optical disc
		[0xB4] = "\U0001F39E", // 🎞 U+1F39E Film frames
		[0xB5] = "\U0001F4F7", // 📷 U+1F4F7 Camera
		[0xB6] = "\U0001F39F", // 🎟 U+1F39F Admission tickets
		[0xB7] = "\U0001F3AC", // 🎬 U+1F3AC Clapper board
		[0xB8] = "\U0001F4FD", // 📽 U+1F4FD Film projector
		[0xB9] = "\U0001F4F9", // 📹 U+1F4F9 Video camera
		[0xBA] = "\U0001F4FE", // 📾 U+1F4FE Portable stereo
		[0xBB] = "\U0001F4FB", // 📻 U+1F4FB Radio
		[0xBC] = "\U0001F39A", // 🎚 U+1F39A Level slider
		[0xBD] = "\U0001F39B", // 🎛 U+1F39B Control knobs
		[0xBE] = "\U0001F4FA", // 📺 U+1F4FA Television
		[0xBF] = "\U0001F4BB", // 💻 U+1F4BB Personal computer
		[0xC0] = "\U0001F5A5", // 🖥 U+1F5A5 Desktop computer
		[0xC1] = "\U0001F5A6", // 🖦 U+1F5A6 Keyboard and mouse
		[0xC2] = "\U0001F5A7", // 🖧 U+1F5A7 Three networked computers
		[0xC3] = "\U0001F579", // 🕹 U+1F579 Joystick
		[0xC4] = "\U0001F3AE", // 🎮 U+1F3AE Video game
		[0xC5] = "\U0001F57B", // 🕻 U+1F57B Left hand telephone receiver
		[0xC6] = "\U0001F57C", // 🕼 U+1F57C Telephone receiver with page
		[0xC7] = "\U0001F4DF", // 📟 U+1F4DF Pager
		[0xC8] = "\U0001F581", // 🖁 U+1F581 Clamshell mobile phone
		[0xC9] = "\U0001F580", // 🖀 U+1F580 Telephone on top of modem
		[0xCA] = "\U0001F5A8", // 🖨 U+1F5A8 Printer
		[0xCB] = "\U0001F5A9", // 🖩 U+1F5A9 Pocket calculator
		[0xCC] = "\U0001F5BF", // 🖿 U+1F5BF Black folder
		[0xCD] = "\U0001F5AA", // 🖪 U+1F5AA Black hard shell floppy disk
		[0xCE] = "\U0001F5DC", // 🗜 U+1F5DC Compression
		[0xCF] = "\U0001F512", // 🔒 U+1F512 Lock
		[0xD0] = "\U0001F513", // 🔓 U+1F513 Open lock
		[0xD1] = "\U0001F5DD", // 🗝 U+1F5DD Old key
		[0xD2] = "\U0001F4E5", // 📥 U+1F4E5 Inbox tray
		[0xD3] = "\U0001F4E4", // 📤 U+1F4E4 Outbox tray
		[0xD4] = "\U0001F573", // 🕳 U+1F573 Hole
		[0xD5] = "\U0001F323", // 🌣 U+1F323 White sun
		[0xD6] = "\U0001F324", // 🌤 U+1F324 White sun with small cloud
		[0xD7] = "\U0001F325", // 🌥 U+1F325 White sun behind cloud
		[0xD8] = "\U0001F326", // 🌦 U+1F326 White sun behind cloud with rain
		[0xD9] = "\u2601", // ☁ U+2601 Cloud
		[0xDA] = "\U0001F327", // 🌧 U+1F327 Cloud with rain
		[0xDB] = "\U0001F328", // 🌨 U+1F328 Cloud with snow
		[0xDC] = "\U0001F329", // 🌩 U+1F329 Cloud with lightning
		[0xDD] = "\U0001F32A", // 🌪 U+1F32A Cloud with tornado
		[0xDE] = "\U0001F32C", // 🌬 U+1F32C Wind blowing face
		[0xDF] = "\U0001F32B", // 🌫 U+1F32B Fog
		[0xE0] = "\U0001F31C", // 🌜 U+1F31C Last quarter moon with face
		[0xE1] = "\U0001F321", // 🌡 U+1F321 Thermometer
		[0xE2] = "\U0001F6CB", // 🛋 U+1F6CB Couch and lamp
		[0xE3] = "\U0001F6CF", // 🛏 U+1F6CF Bed
		[0xE4] = "\U0001F37D", // 🍽 U+1F37D Fork and knife with plate
		[0xE5] = "\U0001F378", // 🍸 U+1F378 Cocktail glass
		[0xE6] = "\U0001F6CE", // 🛎 U+1F6CE Bellhop bell
		[0xE7] = "\U0001F6CD", // 🛍 U+1F6CD Shopping bags
		[0xE8] = "\u24C5", // Ⓟ U+24C5 Circled latin capital letter P
		[0xE9] = "\u267F", // ♿ U+267F Wheelchair symbol
		[0xEA] = "\U0001F6C6", // 🛆 U+1F6C6 Triangle with rounded corners
		[0xEB] = "\U0001F588", // 🖈 U+1F588 Black pushpin
		[0xEC] = "\U0001F393", // 🎓 U+1F393 Graduation cap
		[0xED] = "\U0001F5E4", // 🗤 U+1F5E4 Three rays above
		[0xEE] = "\U0001F5E5", // 🗥 U+1F5E5 Three rays below
		[0xEF] = "\U0001F5E6", // 🗦 U+1F5E6 Three rays left
		[0xF0] = "\U0001F5E7", // 🗧 U+1F5E7 Three rays right
		[0xF1] = "\U0001F6EA", // 🛪 U+1F6EA Northeast-pointing airplane
		[0xF2] = "\U0001F43F", // 🐿 U+1F43F Chipmunk
		[0xF3] = "\U0001F426", // 🐦 U+1F426 Bird
		[0xF4] = "\U0001F41F", // 🐟 U+1F41F Fish
		[0xF5] = "\U0001F415", // 🐕 U+1F415 Dog
		[0xF6] = "\U0001F408", // 🐈 U+1F408 Cat
		[0xF7] = "\U0001F66C", // 🙬 U+1F66C Leftwards rocket
		[0xF8] = "\U0001F66E", // 🙮 U+1F66E Rightwards rocket
		[0xF9] = "\U0001F66D", // 🙭 U+1F66D Upwards rocket
		[0xFA] = "\U0001F66F", // 🙯 U+1F66F Downwards rocket
		[0xFB] = "\U0001F5FA", // 🗺 U+1F5FA World map
		[0xFC] = "\U0001F30D", // 🌍 U+1F30D Earth globe Europe-Africa
		[0xFD] = "\U0001F30F", // 🌏 U+1F30F Earth globe Asia-Australia
		[0xFE] = "\U0001F30E", // 🌎 U+1F30E Earth globe Americas
		[0xFF] = "\U0001F54A", // 🕊 U+1F54A Dove of peace
	};

	public static string? ToUnicode(byte symbolCode)
	{
		return Table[symbolCode];
	}

	public static byte[]? ToUtf8Bytes(byte symbolCode)
	{
		var s = ToUnicode(symbolCode);
		return s is null ? null : Encoding.UTF8.GetBytes(s);
	}
}
