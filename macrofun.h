// xll_macrofun.h - tools for automating spreadsheet creation
// This header makes it convenient to call macro functions 
// documented in https://xlladdins.github.io/Excel4Macros/
#pragma once
#include "excel.h"

#define XLL_RGB_COLOR(X) \
X(WHITE, 0xFFFFFF, ({255, 255, 255})) \
X(RED, 0xFF0000, ({255, 0, 0})) \
X(MAGENTA, 0xFF00FF, ({255, 0, 255})) \
X(SILVER, 0xC0C0C0, ({192, 192, 192})) \
X(GRAY, 0x808080, ({128, 128, 128})) \
X(MAROON, 0x800000, ({128, 0, 0})) \
X(DARK_RED, 0x8B0000, ({139, 0, 0})) \
X(BROWN, 0xA52A2A, ({165, 42, 42})) \
X(FIREBRICK, 0xB22222, ({178, 34, 34})) \
X(CRIMSON, 0xDC143C, ({220, 20, 60})) \
X(TOMATO, 0xFF6347, ({255, 99, 71})) \
X(CORAL, 0xFF7F50, ({255, 127, 80})) \
X(INDIAN_RED, 0xCD5C5C, ({205, 92, 92})) \
X(LIGHT_CORAL, 0xF08080, ({240, 128, 128})) \
X(DARK_SALMON, 0xE9967A, ({233, 150, 122})) \
X(SALMON, 0xFA8072, ({250, 128, 114})) \
X(LIGHT_SALMON, 0xFFA07A, ({255, 160, 122})) \
X(ORANGE_RED, 0xFF4500, ({255, 69, 0})) \
X(DARK_ORANGE, 0xFF8C00, ({255, 140, 0})) \
X(ORANGE, 0xFFA500, ({255, 165, 0})) \
X(GOLD, 0xFFD700, ({255, 215, 0})) \
X(DARK_GOLDEN_ROD, 0xB8860B, ({184, 134, 11})) \
X(GOLDEN_ROD, 0xDAA520, ({218, 165, 32})) \
X(PALE_GOLDEN_ROD, 0xEEE8AA, ({238, 232, 170})) \
X(DARK_KHAKI, 0xBDB76B, ({189, 183, 107})) \
X(KHAKI, 0xF0E68C, ({240, 230, 140})) \
X(OLIVE, 0x808000, ({128, 128, 0})) \
X(YELLOW, 0xFFFF00, ({255, 255, 0})) \
X(YELLOW_GREEN, 0x9ACD32, ({154, 205, 50})) \
X(DARK_OLIVE_GREEN, 0x556B2F, ({85, 107, 47})) \
X(OLIVE_DRAB, 0x6B8E23, ({107, 142, 35})) \
X(LAWN_GREEN, 0x7CFC00, ({124, 252, 0})) \
X(CHARTREUSE, 0x7FFF00, ({127, 255, 0})) \
X(GREEN_YELLOW, 0xADFF2F, ({173, 255, 47})) \
X(DARK_GREEN, 0x006400, ({0, 100, 0})) \
X(GREEN, 0x008000, ({0, 128, 0})) \
X(FOREST_GREEN, 0x228B22, ({34, 139, 34})) \
X(LIME, 0x00FF00, ({0, 255, 0})) \
X(LIME_GREEN, 0x32CD32, ({50, 205, 50})) \
X(LIGHT_GREEN, 0x90EE90, ({144, 238, 144})) \
X(PALE_GREEN, 0x98FB98, ({152, 251, 152})) \
X(DARK_SEA_GREEN, 0x8FBC8F, ({143, 188, 143})) \
X(MEDIUM_SPRING_GREEN, 0x00FA9A, ({0, 250, 154})) \
X(SPRING_GREEN, 0x00FF7F, ({0, 255, 127})) \
X(SEA_GREEN, 0x2E8B57, ({46, 139, 87})) \
X(MEDIUM_AQUA_MARINE, 0x66CDAA, ({102, 205, 170})) \
X(MEDIUM_SEA_GREEN, 0x3CB371, ({60, 179, 113})) \
X(LIGHT_SEA_GREEN, 0x20B2AA, ({32, 178, 170})) \
X(DARK_SLATE_GRAY, 0x2F4F4F, ({47, 79, 79})) \
X(TEAL, 0x008080, ({0, 128, 128})) \
X(DARK_CYAN, 0x008B8B, ({0, 139, 139})) \
X(AQUA, 0x00FFFF, ({0, 255, 255})) \
X(CYAN, 0x00FFFF, ({0, 255, 255})) \
X(LIGHT_CYAN, 0xE0FFFF, ({224, 255, 255})) \
X(DARK_TURQUOISE, 0x00CED1, ({0, 206, 209})) \
X(TURQUOISE, 0x40E0D0, ({64, 224, 208})) \
X(MEDIUM_TURQUOISE, 0x48D1CC, ({72, 209, 204})) \
X(PALE_TURQUOISE, 0xAFEEEE, ({175, 238, 238})) \
X(AQUA_MARINE, 0x7FFFD4, ({127, 255, 212})) \
X(POWDER_BLUE, 0xB0E0E6, ({176, 224, 230})) \
X(CADET_BLUE, 0x5F9EA0, ({95, 158, 160})) \
X(STEEL_BLUE, 0x4682B4, ({70, 130, 180})) \
X(CORN_FLOWER_BLUE, 0x6495ED, ({100, 149, 237})) \
X(DEEP_SKY_BLUE, 0x00BFFF, ({0, 191, 255})) \
X(DODGER_BLUE, 0x1E90FF, ({30, 144, 255})) \
X(LIGHT_BLUE, 0xADD8E6, ({173, 216, 230})) \
X(SKY_BLUE, 0x87CEEB, ({135, 206, 235})) \
X(LIGHT_SKY_BLUE, 0x87CEFA, ({135, 206, 250})) \
X(MIDNIGHT_BLUE, 0x191970, ({25, 25, 112})) \
X(NAVY, 0x000080, ({0, 0, 128})) \
X(DARK_BLUE, 0x00008B, ({0, 0, 139})) \
X(MEDIUM_BLUE, 0x0000CD, ({0, 0, 205})) \
X(BLUE, 0x0000FF, ({0, 0, 255})) \
X(ROYAL_BLUE, 0x4169E1, ({65, 105, 225})) \
X(BLUE_VIOLET, 0x8A2BE2, ({138, 43, 226})) \
X(INDIGO, 0x4B0082, ({75, 0, 130})) \
X(DARK_SLATE_BLUE, 0x483D8B, ({72, 61, 139})) \
X(SLATE_BLUE, 0x6A5ACD, ({106, 90, 205})) \
X(MEDIUM_SLATE_BLUE, 0x7B68EE, ({123, 104, 238})) \
X(MEDIUM_PURPLE, 0x9370DB, ({147, 112, 219})) \
X(DARK_MAGENTA, 0x8B008B, ({139, 0, 139})) \
X(DARK_VIOLET, 0x9400D3, ({148, 0, 211})) \
X(DARK_ORCHID, 0x9932CC, ({153, 50, 204})) \
X(MEDIUM_ORCHID, 0xBA55D3, ({186, 85, 211})) \
X(PURPLE, 0x800080, ({128, 0, 128})) \
X(THISTLE, 0xD8BFD8, ({216, 191, 216})) \
X(PLUM, 0xDDA0DD, ({221, 160, 221})) \
X(VIOLET, 0xEE82EE, ({238, 130, 238})) \
X(FUCHSIA, 0xFF00FF, ({255, 0, 255})) \
X(ORCHID, 0xDA70D6, ({218, 112, 214})) \
X(MEDIUM_VIOLET_RED, 0xC71585, ({199, 21, 133})) \
X(PALE_VIOLET_RED, 0xDB7093, ({219, 112, 147})) \
X(DEEP_PINK, 0xFF1493, ({255, 20, 147})) \
X(HOT_PINK, 0xFF69B4, ({255, 105, 180})) \
X(LIGHT_PINK, 0xFFB6C1, ({255, 182, 193})) \
X(PINK, 0xFFC0CB, ({255, 192, 203})) \
X(ANTIQUE_WHITE, 0xFAEBD7, ({250, 235, 215})) \
X(BEIGE, 0xF5F5DC, ({245, 245, 220})) \
X(BISQUE, 0xFFE4C4, ({255, 228, 196})) \
X(BLANCHED_ALMOND, 0xFFEBCD, ({255, 235, 205})) \
X(WHEAT, 0xF5DEB3, ({245, 222, 179})) \
X(CORN_SILK, 0xFFF8DC, ({255, 248, 220})) \
X(LEMON_CHIFFON, 0xFFFACD, ({255, 250, 205})) \
X(LIGHT_GOLDEN_ROD_YELLOW, 0xFAFAD2, ({250, 250, 210})) \
X(LIGHT_YELLOW, 0xFFFFE0, ({255, 255, 224})) \
X(SADDLE_BROWN, 0x8B4513, ({139, 69, 19})) \
X(SIENNA, 0xA0522D, ({160, 82, 45})) \
X(CHOCOLATE, 0xD2691E, ({210, 105, 30})) \
X(PERU, 0xCD853F, ({205, 133, 63})) \
X(SANDY_BROWN, 0xF4A460, ({244, 164, 96})) \
X(BURLY_WOOD, 0xDEB887, ({222, 184, 135})) \
X(TAN, 0xD2B48C, ({210, 180, 140})) \
X(ROSY_BROWN, 0xBC8F8F, ({188, 143, 143})) \
X(MOCCASIN, 0xFFE4B5, ({255, 228, 181})) \
X(NAVAJO_WHITE, 0xFFDEAD, ({255, 222, 173})) \
X(PEACH_PUFF, 0xFFDAB9, ({255, 218, 185})) \
X(MISTY_ROSE, 0xFFE4E1, ({255, 228, 225})) \
X(LAVENDER_BLUSH, 0xFFF0F5, ({255, 240, 245})) \
X(LINEN, 0xFAF0E6, ({250, 240, 230})) \
X(OLD_LACE, 0xFDF5E6, ({253, 245, 230})) \
X(PAPAYA_WHIP, 0xFFEFD5, ({255, 239, 213})) \
X(SEA_SHELL, 0xFFF5EE, ({255, 245, 238})) \
X(MINT_CREAM, 0xF5FFFA, ({245, 255, 250})) \
X(SLATE_GRAY, 0x708090, ({112, 128, 144})) \
X(LIGHT_SLATE_GRAY, 0x778899, ({119, 136, 153})) \
X(LIGHT_STEEL_BLUE, 0xB0C4DE, ({176, 196, 222})) \
X(LAVENDER, 0xE6E6FA, ({230, 230, 250})) \
X(FLORAL_WHITE, 0xFFFAF0, ({255, 250, 240})) \
X(ALICE_BLUE, 0xF0F8FF, ({240, 248, 255})) \
X(GHOST_WHITE, 0xF8F8FF, ({248, 248, 255})) \
X(HONEYDEW, 0xF0FFF0, ({240, 255, 240})) \
X(IVORY, 0xFFFFF0, ({255, 255, 240})) \
X(AZURE, 0xF0FFFF, ({240, 255, 255})) \
X(SNOW, 0xFFFAFA, ({255, 250, 250})) \
X(BLACK, 0x000000, ({0, 0, 0})) \
X(DIM_GRAY, 0x696969, ({105, 105, 105})) \
X(DARK_GRAY, 0xA9A9A9, ({169, 169, 169})) \
X(LIGHT_GRAY, 0xD3D3D3, ({211, 211, 211})) \
X(GAINSBORO, 0xDCDCDC, ({220, 220, 220})) \
X(WHITE_SMOKE, 0xF5F5F5, ({245, 245, 245})) \
X(DIM_GREY, 0x696969, ({105, 105, 105})) \
X(GREY, 0x808080, ({128, 128, 128})) \
X(DARK_GREY, 0xA9A9A9, ({169, 169, 169})) \
X(LIGHT_GREY, 0xD3D3D3, ({211, 211, 211})) \
X(MS_RED, 0xF65314, ({246, 83, 20})) \
X(MS_GREEN, 0x7CBB00, ({124, 187, 0})) \
X(MS_BLUE, 0x00A1F1, ({0, 161, 241})) \
X(MS_ORANGE, 0xFFBB00, ({255, 187, 0})) \

#define XLL_RGB_COLOR_ENUM(name, rgb, rgbv) \
constexpr unsigned long XLL_RGB_COLOR_##name = rgb;
XLL_RGB_COLOR(XLL_RGB_COLOR_ENUM)
#undef XLL_RGB_COLOR_ENUM

namespace xll {

	struct Alignment {
		OPER horiz_align = Missing;   // Horizontal alignment
		OPER wrap = Missing;          // boolean
		OPER vert_align = Missing;    // Vertical alignment
		OPER orientation = Missing;   // Text rotation
		OPER add_indent = Missing;    // boolean
		OPER shrink_to_fit = Missing; // boolean
		OPER merge_cells = Missing;   // boolean

		~Alignment()
		{
			Excel(xlcAlignment, horiz_align, wrap, vert_align, orientation, add_indent, shrink_to_fit, merge_cells);
		}

		enum class Horizontal {
			General = 1,
			Left = 2,
			Center = 3,
			Right = 4,
			Fill = 5,
			Justify = 6,
			CenterAcrossSelection = 7
		};
		enum class Vertical {
			Top = 1,
			Center = 2,
			Bottom = 3,
			Justify = 4
		};
		enum class Orientation {
			Horizontal = 0,
			Vertical = 1,
			Upward = 2,
			Downward = 3,
			Automatic = 4
		};
	};
	void AlignHorizontal(Alignment::Horizontal align)
	{
		Alignment a{ .horiz_align = OPER(static_cast<int>(align)) };
	}
	void AlignVertical(Alignment::Vertical align)
	{
		Alignment a{ .vert_align = OPER(static_cast<int>(align)) };
	}
	void AlignOrientation(Alignment::Orientation align)
	{
		Alignment a{ .orientation = OPER(static_cast<int>(align)) };
	}
	void AlignWrap(bool b = true)
	{
		Alignment a{ .wrap = OPER(b) };
	}
	void AlignShrink(bool b = true)
	{
		Alignment a{ .shrink_to_fit = OPER(b) };
	}
	void AlignMerge(bool b = true)
	{
		Alignment a{ .merge_cells = OPER(b) };
	}

	// https://xlladdins.github.io/Excel4Macros/edit.color.html
	// Equivalent to clicking the Modify button on the Color tab, 
	// which appears when you click the Options command on the Tools menu.
	// Defines the color for one of the 56 color palette boxes.
	class EditColor {
		inline static int count = 56; // max color palette index
	public:
		int index;
		unsigned char r, g, b; // Enforces 0 <= r, g, b <= 255.
		EditColor(unsigned char r, unsigned char g, unsigned char b, int index = count)
			: index(index), r(r), g(g), b(b)
		{ 
			--count;
			ensure(count > 0);
			ensure(Excel(xlcEditColor, index, r, g, b));
		}
		EditColor(unsigned long rgb)
			: EditColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF)
		{ }
		operator int() const
		{
			return index;
		}
	};

	// FORMAT.FONT(name_text, size_num, bold, italic, underline, strike, color, outline, shadow)
// https://xlladdins.github.io/Excel4Macros/format.font.html
	struct FormatFont {
		OPER name_text = Missing; // font name, e.g., "Calibri"
		OPER size_num = Missing;  // font size in points
		OPER bold = Missing;      // boolean
		OPER italic = Missing;    // boolean
		OPER underline = Missing; // boolean
		OPER strike = Missing;    // boolean
		OPER color = Missing;     // color palette index
		OPER outline = Missing;   // boolean
		OPER shadow = Missing;    // boolean

		~FormatFont()
		{
			Excel(xlcFormatFont, name_text, size_num,
				bold, italic, underline, strike,
				color, outline, shadow);
		}
	};
	void FormatFontName(const OPER& name)
	{
		FormatFont ff{ .name_text = name };
	}
	void FormatFontSize(int points)
	{
		FormatFont ff{ .size_num = OPER(points) };
	}
	void FormatFontBold(bool b = true)
	{
		FormatFont ff{ .bold = OPER(b) };
	}
	void FormatFontItalic(bool b = true)
	{
		FormatFont ff{ .italic = OPER(b) };
	}
	void FormatFontUnderline(bool b = true)
	{
		FormatFont ff{ .underline = OPER(b) };
	}
	void FormatFontStrike(bool b = true)
	{
		FormatFont ff{ .strike = OPER(b) };
	}
	void FormatFontColor(int index)
	{
		FormatFont ff{ .color = OPER(index) };
	}


} // namespace xll
