// type.h - Type s slowly into the active cell.
#pragma once
#include "xll.h"

void DoEvents(int ms = 0);

namespace xll {

	// Pause in milliseconds between typing characters and waiting after typing.
	inline void Type(const char* s, int pause, int wait)
	{
		OPER ac = Excel(xlfActiveCell);
		OPER ct = Excel(xlfGetDocument, 14); // calculation type
		Excel(xlcOptionsCalculation, 3); // manual calc

		OPER t(s);
		size_t n = strlen(s);

		for (size_t i = 1; i <= n; ++i) {
			t.val.str[0] = static_cast<char>(i);
			Excel(xlSet, ac, t);
			if (Excel(xlAbort)) {
				return;
			}
			// http://blogs.msdn.com/b/excel/archive/2006/06/30/642465.aspx
			DoEvents(pause); // prevent macro from making Excel unresponsive.
		}
		DoEvents(wait);
		Excel(xlcFormula, OPER(s), ac);
		Excel(xlcOptionsCalculation, ct);
	}
} // namespace xll