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
		unsigned char r, g, b; // Enforces 0 <= r, g, b <= 255.
		EditColor(unsigned char r, unsigned char g, unsigned char b)
			: r(r), g(g), b(b)
		{ }
		// Add from end of color palette
		OPER Set(int index = 0) const
		{
			if (index == 0) {
				ensure(count > 1);
				index = --count;
			}
			
			return Excel(xlcEditColor, index, r, g, b);
		}
		// Microsoft color palette
		static OPER MSred()
		{
			return EditColor(0xF6, 0x53, 0x14).Set();
		}
		static OPER MSgreen()
		{
			return EditColor(0x7C, 0xBB, 0x00).Set();
		}
		static OPER MSblue()
		{
			return EditColor(0x00, 0xA1, 0xF1).Set();
		}
		static OPER MSorange()
		{
			return EditColor(0, 0, 0).Set();
		}
	};

} // namespace xll
