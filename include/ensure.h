// ensure.h - assert replacement that throws instead of calling abort()
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
// #define NENSURE before including to turn ensure checking off
// #define DEBUG_BREAK to turn on debug break on ensure failure
#pragma once
//#include <debugging> // https://en.cppreference.com/w/cpp/header/debugging

#include <stdexcept>
#include <string>

#ifndef ensure

// Define NENSURE to turn off ensure.
#ifdef NENSURE
#define ensure(x) x
#else
#ifdef _WIN32
#include <Windows.h>
#define ENSURE_DEBUG_BREAK() DebugBreak()
#else
#define ENSURE_DEBUG_BREAK() __builtin_debugtrap();
#endif

#define ENSURE_HASH_(x) #x
#define ENSURE_STRZ_(x) ENSURE_HASH_(x)
#define ENSURE_FILE "file: " __FILE__
#ifdef __FUNCTION__
#define ENSURE_FUNC "\nfunction: " __FUNCTION__
#else
#define ENSURE_FUNC ""
#endif
#define ENSURE_LINE "\nline: " ENSURE_STRZ_(__LINE__)
#define ENSURE_SPOT ENSURE_FILE ENSURE_LINE ENSURE_FUNC

#if defined(_DEBUG) && defined(DEBUG_BREAK)
#define ensure(e) if (!(e)) { ENSURE_DEBUG_BREAK(); } else (void)0
#define ensure_message(e, s) if (!(e)) { ENSURE_DEBUG_BREAK(); }else (void)0
#else
#define ensure(e) if (!(e)) { \
		throw std::runtime_error(ENSURE_SPOT "\nensure: \"" #e "\" failed"); \
		} else (void)0
#define ensure_message(e, s) if (!(e)) { \
		throw std::runtime_error(std::string(ENSURE_SPOT "\nensure: \"" #e "\"\nmessage: ") + s); \
		} else (void)0
#endif // _DEBUG

#endif // NENSURE
#endif // ensure
