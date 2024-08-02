// on.h - xlcOnXXX functions
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
// https://xlladdins.github.io/Excel4Macros/#O
#pragma once
#include "auto.h"
#include "excel.h"

// For use with On<Key>
#define ON_SHIFT       "+"
#define ON_CTRL        "^"
#define ON_ALT         "%"
#define ON_COMMAND     "*"
#define ON_ENTER       "~"
#define ON_BACKSPACE   "{BACKSPACE}" 
#define ON_BS          "{BS}" 
#define ON_BREAK       "{BREAK}" 
#define ON_CAPS_LOCK   "{CAPSLOCK}" 
#define ON_CLEAR       "{CLEAR}" 
#define ON_DELETE      "{DELETE}" 
#define ON_DEL         "{DEL}" 
#define ON_DOWN_ARROW  "{DOWN}" 
#define ON_END         "{END}" 
#define ON_ENTER_NUM   "{ENTER}" 
#define ON_ESCAPE      "{ESCAPE}" 
#define ON_ESC         "{ESC}" 
#define ON_F1          "{F1}" 
#define ON_F2          "{F2}" 
#define ON_F3          "{F3}" 
#define ON_F4          "{F4}" 
#define ON_F5          "{F5}" 
#define ON_F6          "{F6}" 
#define ON_F7          "{F7}" 
#define ON_F8          "{F8}" 
#define ON_F9          "{F9}" 
#define ON_F10         "{F10}" 
#define ON_F11         "{F11}" 
#define ON_F12         "{F12}" 
#define ON_F13         "{F13}" 
#define ON_F14         "{F14}" 
#define ON_F15         "{F15}" 
#define ON_HELP        "{HELP}" 
#define ON_HOME        "{HOME}" 
#define ON_INSERT      "{INSERT}" 
#define ON_LEFT_ARROW  "{LEFT}" 
#define ON_NUM_LOCK    "{NUMLOCK}" 
#define ON_PAGE_DOWN   "{PGDN}" 
#define ON_PAGE_UP     "{PGUP}" 
#define ON_RETURN      "{RETURN}" 
#define ON_RIGHT_ARROW "{RIGHT}" 
#define ON_SCROLL_LOCK "{SCROLLLOCK}" 
#define ON_TAB         "{TAB}" 
#define ON_UP_ARROW    "{UP}" 

namespace xll {

	// Calls on xlcOnXXX but overwrites existing macros.
	// TODO: ???Keep a list of all On macros???
	template<int Key>
	class On {
		using cstr = const char*;
	public:
		On(cstr text, cstr macro)
		{
			Auto<OpenAfter> xao([text, macro]() {
				return Excel(Key, OPER(text), OPER(macro)) == true;
				});
			///* TODO: This seems to cause memory leaks!!!
			Auto<CloseBefore> xac([text]() {
				return Excel(Key, OPER(text)) == true;
			});
			//*/
		}
		On(cstr text, cstr macro, bool activate)
		{
			static_assert(Key == xlcOnSheet);

			Auto<OpenAfter> xao([text, macro, activate]() {
				return Excel(Key, OPER(text), OPER(macro), OPER(activate)) == true;
				});
			Auto<CloseBefore> xac([text]() {
				return Excel(Key, OPER(text)) == true;
				});
		}
		On(const OPER& time, cstr macro, const OPER& tolerance, bool insert)
		{
			static_assert(Key == xlcOnTime);

			Auto<OpenAfter> xao([time, macro, tolerance, insert]() {
				return Excel(Key, OPER(time), OPER(macro), OPER(tolerance), OPER(insert)) == true;
				});
			Auto<CloseBefore> xac([]() {
				return Excel(Key) == true;
				});
		}
	};

	// https://xlladdins.github.io/Excel4Macros/on.data.html
	struct OnData : public On<xlcOnData> {
		OnData(const char* text, const char* macro) 
			: On<xlcOnData>(text, macro) 
		{ } 
	};

	// https://xlladdins.github.io/Excel4Macros/on.doubleclick.html
	struct OnDoulbeclick : public On<xlcOnDoubleclick> {
		OnDoulbeclick(const char* text, const char* macro)
			: On<xlcOnDoubleclick>(text, macro)
		{ }
	};

	// https://xlladdins.github.io/Excel4Macros/on.entry.html
	struct OnEntry : public On<xlcOnEntry> {
		OnEntry(const char* text, const char* macro)
			: On<xlcOnEntry>(text, macro)
		{ }
	};

	// https://xlladdins.github.io/Excel4Macros/on.key.html
	struct OnKey : public On<xlcOnKey> {
		OnKey(const char* text, const char* macro)
			: On<xlcOnKey>(text, macro)
		{ }
	};

	// https://xlladdins.github.io/Excel4Macros/on.recalc.html
	struct OnRecalc : public On<xlcOnRecalc> {
		OnRecalc(const char* text, const char* macro)
			: On<xlcOnRecalc>(text, macro)
		{ }
	};

	// https://xlladdins.github.io/Excel4Macros/on.sheet.html
	struct OnSheet : public On<xlcOnSheet> {
		OnSheet(const char* text, const char* macro, bool activate)
			: On<xlcOnSheet>(text, macro, activate)
		{ }
	};

	// https://xlladdins.github.io/Excel4Macros/on.time.html
	struct OnTime : public On<xlcOnTime> {
		OnTime(const OPER& time, const char* macro, const OPER& tolerance, bool insert)
			: On<xlcOnTime>(time, macro, tolerance, insert)
		{ }
	};

	// https://xlladdins.github.io/Excel4Macros/on.window.html
	struct OnWindow : public On<xlcOnWindow> {
		OnWindow(const char* text, const char* macro)
			: On<xlcOnWindow>(text, macro)
		{ }
	};

} // namespace xll