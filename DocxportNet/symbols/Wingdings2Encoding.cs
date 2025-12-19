using System.Text;

namespace DocxportNet.symbols;

public static class Wingdings2Map
{
	// Key = Wingdings 2 code (byte 0x00..0xFF as used by the Wingdings 2 font)
	// Value = Unicode string (encode with UTF-8 as needed)
	public static readonly Dictionary<byte, string> Wingdings2ToUnicode = new() {
		[0x20] = "\u0020",       // Space U+0020 :contentReference[oaicite:1]{index=1}
		[0x21] = "\U0001F58A",   // 🖊 U+1F58A :contentReference[oaicite:2]{index=2}
		[0x22] = "\U0001F58B",   // 🖋 U+1F58B :contentReference[oaicite:3]{index=3}
		[0x23] = "\U0001F58C",   // 🖌 U+1F58C :contentReference[oaicite:4]{index=4}
		[0x24] = "\U0001F58D",   // 🖍 U+1F58D :contentReference[oaicite:5]{index=5}
		[0x25] = "\u2704",       // ✄ U+2704 :contentReference[oaicite:6]{index=6}
		[0x26] = "\u2700",       // ✀ U+2700 :contentReference[oaicite:7]{index=7}
		[0x27] = "\U0001F57E",   // 🕾 U+1F57E :contentReference[oaicite:8]{index=8}
		[0x28] = "\U0001F57D",   // 🕽 U+1F57D :contentReference[oaicite:9]{index=9}
		[0x29] = "\U0001F5C5",   // 🗅 U+1F5C5 :contentReference[oaicite:10]{index=10}
		[0x2A] = "\U0001F5C6",   // 🗆 U+1F5C6 :contentReference[oaicite:11]{index=11}
		[0x2B] = "\U0001F5C7",   // 🗇 U+1F5C7 :contentReference[oaicite:12]{index=12}
		[0x2C] = "\U0001F5C8",   // 🗈 U+1F5C8 :contentReference[oaicite:13]{index=13}
		[0x2D] = "\U0001F5C9",   // 🗉 U+1F5C9 :contentReference[oaicite:14]{index=14}
		[0x2E] = "\U0001F5CA",   // 🗊 U+1F5CA :contentReference[oaicite:15]{index=15}
		[0x2F] = "\U0001F5CB",   // 🗋 U+1F5CB :contentReference[oaicite:16]{index=16}
		[0x30] = "\U0001F5CC",   // 🗌 U+1F5CC :contentReference[oaicite:17]{index=17}
		[0x31] = "\U0001F5CD",   // 🗍 U+1F5CD :contentReference[oaicite:18]{index=18}
		[0x32] = "\U0001F4CB",   // 📋 U+1F4CB :contentReference[oaicite:19]{index=19}
		[0x33] = "\U0001F5D1",   // 🗑 U+1F5D1 :contentReference[oaicite:20]{index=20}
		[0x34] = "\U0001F5D4",   // 🗔 U+1F5D4 :contentReference[oaicite:21]{index=21}
		[0x35] = "\U0001F5B5",   // 🖵 U+1F5B5 :contentReference[oaicite:22]{index=22}
		[0x36] = "\U0001F5B6",   // 🖶 U+1F5B6 :contentReference[oaicite:23]{index=23}
		[0x37] = "\U0001F5B7",   // 🖷 U+1F5B7 :contentReference[oaicite:24]{index=24}
		[0x38] = "\U0001F5B8",   // 🖸 U+1F5B8 :contentReference[oaicite:25]{index=25}
		[0x39] = "\U0001F5AD",   // 🖭 U+1F5AD :contentReference[oaicite:26]{index=26}
		[0x3A] = "\U0001F5AF",   // 🖯 U+1F5AF :contentReference[oaicite:27]{index=27}
		[0x3B] = "\U0001F5B1",   // 🖱 U+1F5B1 :contentReference[oaicite:28]{index=28}
		[0x3C] = "\U0001F592",   // 🖒 U+1F592 :contentReference[oaicite:29]{index=29}
		[0x3D] = "\U0001F593",   // 🖓 U+1F593 :contentReference[oaicite:30]{index=30}
		[0x3E] = "\U0001F598",   // 🖘 U+1F598 :contentReference[oaicite:31]{index=31}
		[0x3F] = "\U0001F599",   // 🖙 U+1F599 :contentReference[oaicite:32]{index=32}
		[0x40] = "\U0001F59A",   // 🖚 U+1F59A :contentReference[oaicite:33]{index=33}
		[0x41] = "\U0001F59B",   // 🖛 U+1F59B :contentReference[oaicite:34]{index=34}
		[0x42] = "\U0001F448",   // 👈 U+1F448 :contentReference[oaicite:35]{index=35}
		[0x43] = "\U0001F449",   // 👉 U+1F449 :contentReference[oaicite:36]{index=36}
		[0x44] = "\U0001F59C",   // 🖜 U+1F59C :contentReference[oaicite:37]{index=37}
		[0x45] = "\U0001F59D",   // 🖝 U+1F59D :contentReference[oaicite:38]{index=38}
		[0x46] = "\U0001F59E",   // 🖞 U+1F59E :contentReference[oaicite:39]{index=39}
		[0x47] = "\U0001F59F",   // 🖟 U+1F59F :contentReference[oaicite:40]{index=40}
		[0x48] = "\U0001F5A0",   // 🖠 U+1F5A0 :contentReference[oaicite:41]{index=41}
		[0x49] = "\U0001F5A1",   // 🖡 U+1F5A1 :contentReference[oaicite:42]{index=42}
		[0x4A] = "\U0001F446",   // 👆 U+1F446 :contentReference[oaicite:43]{index=43}
		[0x4B] = "\U0001F447",   // 👇 U+1F447 :contentReference[oaicite:44]{index=44}
		[0x4C] = "\U0001F5A2",   // 🖢 U+1F5A2 :contentReference[oaicite:45]{index=45}
		[0x4D] = "\U0001F5A3",   // 🖣 U+1F5A3 :contentReference[oaicite:46]{index=46}
		[0x4E] = "\U0001F591",   // 🖑 U+1F591 :contentReference[oaicite:47]{index=47}
		[0x4F] = "\U0001F5F4",   // 🗴 U+1F5F4 :contentReference[oaicite:48]{index=48}
		[0x50] = "\u2713",       // ✓ U+2713 :contentReference[oaicite:49]{index=49}
		[0x51] = "\U0001F5F5",   // 🗵 U+1F5F5 :contentReference[oaicite:50]{index=50}
		[0x52] = "\u2611",       // ☑ U+2611 :contentReference[oaicite:51]{index=51}
		[0x53] = "\u2612",       // ☒ U+2612 :contentReference[oaicite:52]{index=52}
		[0x54] = "\u2612",       // ☒ U+2612 (bold style) :contentReference[oaicite:53]{index=53}
		[0x55] = "\u2BBE",       // ⮾ U+2BBE :contentReference[oaicite:54]{index=54}
		[0x56] = "\u2BBF",       // ⮿ U+2BBF :contentReference[oaicite:55]{index=55}
		[0x57] = "\u29B8",       // ⦸ U+29B8 :contentReference[oaicite:56]{index=56}
		[0x58] = "\u29B8",       // ⦸ U+29B8 (bold style) :contentReference[oaicite:57]{index=57}
		[0x59] = "\U0001F671",   // 🙱 U+1F671 :contentReference[oaicite:58]{index=58}
		[0x5A] = "\U0001F674",   // 🙴 U+1F674 :contentReference[oaicite:59]{index=59}
		[0x5B] = "\U0001F672",   // 🙲 U+1F672 :contentReference[oaicite:60]{index=60}
		[0x5C] = "\U0001F673",   // 🙳 U+1F673 :contentReference[oaicite:61]{index=61}
		[0x5D] = "\u203D",       // ‽ U+203D :contentReference[oaicite:62]{index=62}
		[0x5E] = "\U0001F679",   // 🙹 U+1F679 :contentReference[oaicite:63]{index=63}
		[0x5F] = "\U0001F67A",   // 🙺 U+1F67A :contentReference[oaicite:64]{index=64}
		[0x60] = "\U0001F67B",   // 🙻 U+1F67B :contentReference[oaicite:65]{index=65}
		[0x61] = "\U0001F666",   // 🙦 U+1F666 :contentReference[oaicite:66]{index=66}
		[0x62] = "\U0001F664",   // 🙤 U+1F664 :contentReference[oaicite:67]{index=67}
		[0x63] = "\U0001F665",   // 🙥 U+1F665 :contentReference[oaicite:68]{index=68}
		[0x64] = "\U0001F667",   // 🙧 U+1F667 :contentReference[oaicite:69]{index=69}
		[0x65] = "\U0001F65A",   // 🙚 U+1F65A :contentReference[oaicite:70]{index=70}
		[0x66] = "\U0001F658",   // 🙘 U+1F658 :contentReference[oaicite:71]{index=71}
		[0x67] = "\U0001F659",   // 🙙 U+1F659 :contentReference[oaicite:72]{index=72}
		[0x68] = "\U0001F65B",   // 🙛 U+1F65B :contentReference[oaicite:73]{index=73}

		[0x69] = "\u24EA",       // ⓪ U+24EA :contentReference[oaicite:74]{index=74}
		[0x6A] = "\u2460",       // ① U+2460 :contentReference[oaicite:75]{index=75}
		[0x6B] = "\u2461",       // ② U+2461 :contentReference[oaicite:76]{index=76}
		[0x6C] = "\u2462",       // ③ U+2462 :contentReference[oaicite:77]{index=77}
		[0x6D] = "\u2463",       // ④ U+2463 :contentReference[oaicite:78]{index=78}
		[0x6E] = "\u2464",       // ⑤ U+2464 :contentReference[oaicite:79]{index=79}
		[0x6F] = "\u2465",       // ⑥ U+2465 :contentReference[oaicite:80]{index=80}
		[0x70] = "\u2466",       // ⑦ U+2466 :contentReference[oaicite:81]{index=81}
		[0x71] = "\u2467",       // ⑧ U+2467 :contentReference[oaicite:82]{index=82}
		[0x72] = "\u2468",       // ⑨ U+2468 :contentReference[oaicite:83]{index=83}
		[0x73] = "\u2469",       // ⑩ U+2469 :contentReference[oaicite:84]{index=84}
		[0x74] = "\u24FF",       // ⓿ U+24FF :contentReference[oaicite:85]{index=85}
		[0x75] = "\u2776",       // ❶ U+2776 :contentReference[oaicite:86]{index=86}
		[0x76] = "\u2777",       // ❷ U+2777 :contentReference[oaicite:87]{index=87}
		[0x77] = "\u2778",       // ❸ U+2778 :contentReference[oaicite:88]{index=88}
		[0x78] = "\u2779",       // ❹ U+2779 :contentReference[oaicite:89]{index=89}
		[0x79] = "\u277A",       // ❺ U+277A :contentReference[oaicite:90]{index=90}
		[0x7A] = "\u277B",       // ❻ U+277B :contentReference[oaicite:91]{index=91}
		[0x7B] = "\u277C",       // ❼ U+277C :contentReference[oaicite:92]{index=92}
		[0x7C] = "\u277D",       // ❽ U+277D :contentReference[oaicite:93]{index=93}
		[0x7D] = "\u277E",       // ❾ U+277E :contentReference[oaicite:94]{index=94}
		[0x7E] = "\u277F",       // ❿ U+277F :contentReference[oaicite:95]{index=95}

		[0x80] = "\u2609",       // ☉ U+2609 :contentReference[oaicite:96]{index=96}
		[0x81] = "\U0001F315",   // 🌕 U+1F315 :contentReference[oaicite:97]{index=97}
		[0x82] = "\u263D",       // ☽ U+263D :contentReference[oaicite:98]{index=98}
		[0x83] = "\u263E",       // ☾ U+263E :contentReference[oaicite:99]{index=99}
		[0x84] = "\u2E3F",       // ⸿ U+2E3F :contentReference[oaicite:100]{index=100}
		[0x85] = "\u271D",       // ✝ U+271D :contentReference[oaicite:101]{index=101}
		[0x86] = "\U0001F547",   // 🕇 U+1F547 :contentReference[oaicite:102]{index=102}

		[0x87] = "\U0001F55C",   // 🕜 U+1F55C :contentReference[oaicite:103]{index=103}
		[0x88] = "\U0001F55D",   // 🕝 U+1F55D :contentReference[oaicite:104]{index=104}
		[0x89] = "\U0001F55E",   // 🕞 U+1F55E :contentReference[oaicite:105]{index=105}
		[0x8A] = "\U0001F55F",   // 🕟 U+1F55F :contentReference[oaicite:106]{index=106}
		[0x8B] = "\U0001F560",   // 🕠 U+1F560 :contentReference[oaicite:107]{index=107}
		[0x8C] = "\U0001F561",   // 🕡 U+1F561 :contentReference[oaicite:108]{index=108}
		[0x8D] = "\U0001F562",   // 🕢 U+1F562 :contentReference[oaicite:109]{index=109}
		[0x8E] = "\U0001F563",   // 🕣 U+1F563 :contentReference[oaicite:110]{index=110}
		[0x8F] = "\U0001F564",   // 🕤 U+1F564 :contentReference[oaicite:111]{index=111}
		[0x90] = "\U0001F565",   // 🕥 U+1F565 :contentReference[oaicite:112]{index=112}
		[0x91] = "\U0001F566",   // 🕦 U+1F566 :contentReference[oaicite:113]{index=113}
		[0x92] = "\U0001F567",   // 🕧 U+1F567 :contentReference[oaicite:114]{index=114}

		[0x93] = "\U0001F668",   // 🙨 U+1F668 :contentReference[oaicite:115]{index=115}
		[0x94] = "\U0001F669",   // 🙩 U+1F669 :contentReference[oaicite:116]{index=116}

		[0x95] = "\u2022",       // • U+2022 :contentReference[oaicite:117]{index=117}
		[0x96] = "\u25CF",       // ● U+25CF :contentReference[oaicite:118]{index=118}
		[0x97] = "\u26AB",       // ⚫ U+26AB :contentReference[oaicite:119]{index=119}
		[0x98] = "\u2B24",       // ⬤ U+2B24 :contentReference[oaicite:120]{index=120}
		[0x99] = "\U0001F785",   // 🞅 U+1F785 :contentReference[oaicite:121]{index=121}
		[0x9A] = "\U0001F786",   // 🞆 U+1F786 :contentReference[oaicite:122]{index=122}
		[0x9B] = "\U0001F787",   // 🞇 U+1F787 :contentReference[oaicite:123]{index=123}
		[0x9C] = "\U0001F788",   // 🞈 U+1F788 :contentReference[oaicite:124]{index=124}
		[0x9D] = "\U0001F78A",   // 🞊 U+1F78A :contentReference[oaicite:125]{index=125}
		[0x9E] = "\u29BF",       // ⦿ U+29BF :contentReference[oaicite:126]{index=126}
		[0x9F] = "\u25FE",       // ◾ U+25FE :contentReference[oaicite:127]{index=127}

		[0xA0] = "\u25A0",       // ■ U+25A0 :contentReference[oaicite:128]{index=128}
		[0xA1] = "\u25FC",       // ◼ U+25FC :contentReference[oaicite:129]{index=129}
		[0xA2] = "\u2B1B",       // ⬛ U+2B1B :contentReference[oaicite:130]{index=130}
		[0xA3] = "\u2B1C",       // ⬜ U+2B1C :contentReference[oaicite:131]{index=131}
		[0xA4] = "\U0001F791",   // 🞑 U+1F791 :contentReference[oaicite:132]{index=132}
		[0xA5] = "\U0001F792",   // 🞒 U+1F792 :contentReference[oaicite:133]{index=133}
		[0xA6] = "\U0001F793",   // 🞓 U+1F793 :contentReference[oaicite:134]{index=134}
		[0xA7] = "\U0001F794",   // 🞔 U+1F794 :contentReference[oaicite:135]{index=135}
		[0xA8] = "\u25A3",       // ▣ U+25A3 :contentReference[oaicite:136]{index=136}
		[0xA9] = "\U0001F795",   // 🞕 U+1F795 :contentReference[oaicite:137]{index=137}
		[0xAA] = "\U0001F796",   // 🞖 U+1F796 :contentReference[oaicite:138]{index=138}
		[0xAB] = "\U0001F797",   // 🞗 U+1F797 :contentReference[oaicite:139]{index=139}
		[0xAC] = "\u2B29",       // ⬩ U+2B29 :contentReference[oaicite:140]{index=140}
		[0xAD] = "\u2B25",       // ⬥ U+2B25 :contentReference[oaicite:141]{index=141}
		[0xAE] = "\u25C6",       // ◆ U+25C6 :contentReference[oaicite:142]{index=142}
		[0xAF] = "\u25C7",       // ◇ U+25C7 :contentReference[oaicite:143]{index=143}

		[0xB0] = "\U0001F79A",   // 🞚 U+1F79A :contentReference[oaicite:144]{index=144}
		[0xB1] = "\u25C8",       // ◈ U+25C8 :contentReference[oaicite:145]{index=145}
		[0xB2] = "\U0001F79B",   // 🞛 U+1F79B :contentReference[oaicite:146]{index=146}
		[0xB3] = "\U0001F79C",   // 🞜 U+1F79C :contentReference[oaicite:147]{index=147}
		[0xB4] = "\U0001F79D",   // 🞝 U+1F79D :contentReference[oaicite:148]{index=148}
		[0xB5] = "\u2B2A",       // ⬪ U+2B2A :contentReference[oaicite:149]{index=149}
		[0xB6] = "\u2B27",       // ⬧ U+2B27 :contentReference[oaicite:150]{index=150}
		[0xB7] = "\u29EB",       // ⧫ U+29EB :contentReference[oaicite:151]{index=151}
		[0xB8] = "\u25CA",       // ◊ U+25CA :contentReference[oaicite:152]{index=152}
		[0xB9] = "\U0001F7A0",   // 🞠 U+1F7A0 :contentReference[oaicite:153]{index=153}
		[0xBA] = "\u25D6",       // ◖ U+25D6 :contentReference[oaicite:154]{index=154}
		[0xBB] = "\u25D7",       // ◗ U+25D7 :contentReference[oaicite:155]{index=155}
		[0xBC] = "\u2BCA",       // ⯊ U+2BCA :contentReference[oaicite:156]{index=156}
		[0xBD] = "\u2BCB",       // ⯋ U+2BCB :contentReference[oaicite:157]{index=157}
		[0xBE] = "\u25FC",       // ◼ U+25FC :contentReference[oaicite:158]{index=158}
		[0xBF] = "\u2B25",       // ⬥ U+2B25 :contentReference[oaicite:159]{index=159}

		[0xC0] = "\u2B1F",       // ⬟ U+2B1F :contentReference[oaicite:160]{index=160}
		[0xC1] = "\u2BC2",       // ⯂ U+2BC2 :contentReference[oaicite:161]{index=161}
		[0xC2] = "\u2B23",       // ⬣ U+2B23 :contentReference[oaicite:162]{index=162}
		[0xC3] = "\u2B22",       // ⬢ U+2B22 :contentReference[oaicite:163]{index=163}
		[0xC4] = "\u2BC3",       // ⯃ U+2BC3 :contentReference[oaicite:164]{index=164}
		[0xC5] = "\u2BC4",       // ⯄ U+2BC4 :contentReference[oaicite:165]{index=165}

		[0xC6] = "\U0001F7A1",   // 🞡 U+1F7A1 :contentReference[oaicite:166]{index=166}
		[0xC7] = "\U0001F7A2",   // 🞢 U+1F7A2 :contentReference[oaicite:167]{index=167}
		[0xC8] = "\U0001F7A3",   // 🞣 U+1F7A3 :contentReference[oaicite:168]{index=168}
		[0xC9] = "\U0001F7A4",   // 🞤 U+1F7A4 :contentReference[oaicite:169]{index=169}
		[0xCA] = "\U0001F7A5",   // 🞥 U+1F7A5 :contentReference[oaicite:170]{index=170}
		[0xCB] = "\U0001F7A6",   // 🞦 U+1F7A6 :contentReference[oaicite:171]{index=171}
		[0xCC] = "\U0001F7A7",   // 🞧 U+1F7A7 :contentReference[oaicite:172]{index=172}

		[0xCD] = "\U0001F7A8",   // 🞨 U+1F7A8 :contentReference[oaicite:173]{index=173}
		[0xCE] = "\U0001F7A9",   // 🞩 U+1F7A9 :contentReference[oaicite:174]{index=174}
		[0xCF] = "\U0001F7AA",   // 🞪 U+1F7AA :contentReference[oaicite:175]{index=175}
		[0xD0] = "\U0001F7AB",   // 🞫 U+1F7AB :contentReference[oaicite:176]{index=176}
		[0xD1] = "\U0001F7AC",   // 🞬 U+1F7AC :contentReference[oaicite:177]{index=177}
		[0xD2] = "\U0001F7AD",   // 🞭 U+1F7AD :contentReference[oaicite:178]{index=178}
		[0xD3] = "\U0001F7AE",   // 🞮 U+1F7AE :contentReference[oaicite:179]{index=179}

		[0xD4] = "\U0001F7AF",   // 🞯 U+1F7AF :contentReference[oaicite:180]{index=180}
		[0xD5] = "\U0001F7B0",   // 🞰 U+1F7B0 :contentReference[oaicite:181]{index=181}
		[0xD6] = "\U0001F7B1",   // 🞱 U+1F7B1 :contentReference[oaicite:182]{index=182}
		[0xD7] = "\U0001F7B2",   // 🞲 U+1F7B2 :contentReference[oaicite:183]{index=183}
		[0xD8] = "\U0001F7B3",   // 🞳 U+1F7B3 :contentReference[oaicite:184]{index=184}
		[0xD9] = "\U0001F7B4",   // 🞴 U+1F7B4 :contentReference[oaicite:185]{index=185}
		[0xDA] = "\U0001F7B5",   // 🞵 U+1F7B5 :contentReference[oaicite:186]{index=186}
		[0xDB] = "\U0001F7B6",   // 🞶 U+1F7B6 :contentReference[oaicite:187]{index=187}
		[0xDC] = "\U0001F7B7",   // 🞷 U+1F7B7 :contentReference[oaicite:188]{index=188}
		[0xDD] = "\U0001F7B8",   // 🞸 U+1F7B8 :contentReference[oaicite:189]{index=189}
		[0xDE] = "\U0001F7B9",   // 🞹 U+1F7B9 :contentReference[oaicite:190]{index=190}
		[0xDF] = "\U0001F7BA",   // 🞺 U+1F7BA :contentReference[oaicite:191]{index=191}

		[0xE0] = "\U0001F7BB",   // 🞻 U+1F7BB :contentReference[oaicite:192]{index=192}
		[0xE1] = "\U0001F7BC",   // 🞼 U+1F7BC :contentReference[oaicite:193]{index=193}
		[0xE2] = "\U0001F7BD",   // 🞽 U+1F7BD :contentReference[oaicite:194]{index=194}
		[0xE3] = "\U0001F7BE",   // 🞾 U+1F7BE :contentReference[oaicite:195]{index=195}
		[0xE4] = "\U0001F7BF",   // 🞿 U+1F7BF :contentReference[oaicite:196]{index=196}
		[0xE5] = "\U0001F7C0",   // 🟀 U+1F7C0 :contentReference[oaicite:197]{index=197}
		[0xE6] = "\U0001F7C2",   // 🟂 U+1F7C2 :contentReference[oaicite:198]{index=198}
		[0xE7] = "\U0001F7C4",   // 🟄 U+1F7C4 :contentReference[oaicite:199]{index=199}
		[0xE8] = "\u2726",       // ✦ U+2726 :contentReference[oaicite:200]{index=200}
		[0xE9] = "\U0001F7C9",   // 🟉 U+1F7C9 :contentReference[oaicite:201]{index=201}
		[0xEA] = "\u2605",       // ★ U+2605 :contentReference[oaicite:202]{index=202}
		[0xEB] = "\u2736",       // ✶ U+2736 :contentReference[oaicite:203]{index=203}
		[0xEC] = "\U0001F7CB",   // 🟋 U+1F7CB :contentReference[oaicite:204]{index=204}
		[0xED] = "\u2737",       // ✷ U+2737 :contentReference[oaicite:205]{index=205}
		[0xEE] = "\U0001F7CF",   // 🟏 U+1F7CF :contentReference[oaicite:206]{index=206}
		[0xEF] = "\U0001F7D2",   // 🟒 U+1F7D2 :contentReference[oaicite:207]{index=207}
		[0xF0] = "\u2739",       // ✹ U+2739 :contentReference[oaicite:208]{index=208}
		[0xF1] = "\U0001F7C3",   // 🟃 U+1F7C3 :contentReference[oaicite:209]{index=209}
		[0xF2] = "\U0001F7C7",   // 🟇 U+1F7C7 :contentReference[oaicite:210]{index=210}
		[0xF3] = "\u272F",       // ✯ U+272F :contentReference[oaicite:211]{index=211}
		[0xF4] = "\U0001F7CD",   // 🟍 U+1F7CD :contentReference[oaicite:212]{index=212}
		[0xF5] = "\U0001F7D4",   // 🟔 U+1F7D4 :contentReference[oaicite:213]{index=213}
		[0xF6] = "\u2BCC",       // ⯌ U+2BCC :contentReference[oaicite:214]{index=214}
		[0xF7] = "\u2BCD",       // ⯍ U+2BCD :contentReference[oaicite:215]{index=215}
		[0xF8] = "\u203B",       // ※ U+203B :contentReference[oaicite:216]{index=216}
		[0xF9] = "\u2042",       // ⁂ U+2042 :contentReference[oaicite:217]{index=217}
	};

	public static string? ToUnicode(byte wingdings2Code)
		=> Wingdings2ToUnicode.TryGetValue(wingdings2Code, out var s) ? s : null;

	public static byte[]? ToUtf8Bytes(byte wingdings2Code)
	{
		var s = ToUnicode(wingdings2Code);
		return s is null ? null : Encoding.UTF8.GetBytes(s);
	}
}
