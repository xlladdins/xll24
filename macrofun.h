// xll_macrofun.h - tools for automating spreadsheet creation
// This header makes it convenient to call macro functions 
// documented in https://xlladdins.github.io/Excel4Macros/
#pragma once
#include "excel.h"

namespace xll {

	// https://xlladdins.github.io/Excel4Macros/edit.color.html
	// Equivalent to clicking the Modify button on the Color tab, 
	// which appears when you click the Options command on the Tools menu.
	// Defines the color for one of the 56 color palette boxes.
	class EditColor {
		inline static int count = 56; // color palette index
	public:
		unsigned char r, g, b; // Enforce 0 <= r, g, b <= 255.
		EditColor(unsigned char r, unsigned char g, unsigned char b)
			: r(r), g(g), b(b)
		{ }
		// Add from end of color palette
		int Set(int index = 0) const
		{
			if (index == 0) {
				ensure(count > 1);
				index = --count;
			}
			ensure(Excel(xlcEditColor, index, r, g, b));

			return count;
		}
		// Microsoft color palette
		static OPER MSred()
		{
			EditColor _(0xF6, 0x53, 0x14); return _.Set();
		}
		static int MSgreen()
		{
			EditColor _(0x7C, 0xBB, 0x00); return _.Set();
		}
		static int MSblue()
		{
			EditColor _(0x00, 0xA1, 0xF1); return _.Set();
		}
		static int MSorange()
		{
			EditColor _(0xFF, 0xBB, 0x00); return _.Set();
		}
	};

} // namespace xll
