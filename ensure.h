// ensure.h - assert replacement that throws instead of calling abort()
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// #define NENSURE before including to turn ensure checking off
#pragma once
#include <stdexcept>
#include <string>

// Define NENSURE to turn off ensure.
#ifdef NENSURE

#define ensure(x)

#else

#define ENSURE_HASH_(x) #x
#define ENSURE_STRZ_(x) ENSURE_HASH_(x)
#define ENSURE_FILE "file: " __FILE__
#ifdef __FUNCTION__
#define ENSURE_FUNC "\nfunction: " __FUNCTION__
#else
#define ENSURE_FUNC ""
#endif // __FUNCTION__
#define ENSURE_LINE "\nline: " ENSURE_STRZ_(__LINE__)
#define ENSURE_SPOT ENSURE_FILE ENSURE_LINE ENSURE_FUNC

#if defined(DEBUG_BREAK)
#define ensure(e) if (!(e)) { DebugBreak(); }
#define ensure_message(e, s) if (!(e)) { DebugBreak(); }
#else
#define ensure(e) if (!(e)) { \
		throw std::runtime_error(ENSURE_SPOT "\nensure: \"" #e "\""); \
		} else (void)0;
#define ensure_message(e, s) if (!(e)) { \
		throw std::runtime_error(std::string(ENSURE_SPOT "\nensure: \"" #e "\"\nmessage: ") + s); \
		} else (void)0;
#endif // DEBUG_BREAK

#endif // NENSURE

