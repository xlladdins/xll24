cmake_minimum_required(VERSION 3.12)
project(xll
	VERSION 0.1.0
	DESCRIPTION "A library for creating Excel add-ins."
	HOMEPAGE_URL "https://github.com/xlladdins/xll24"
)

set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED True)

add_library(STATIC
src/addin.cpp
src/debug.cpp
src/dllmain.cpp
src/evaluate.cpp
src/fpx.c
src/paste.cpp
src/range.cpp
src/xlauto.cpp
src/XLCALL.CPP
)

target_include_directories(xll PRIVATE ${CMAKE_SOURCE_DIR}/include)

target_compile_options(xll PRIVATE 
$<$<CONFIG:DEBUG>:/Zi /JMC /nologo /W4 /WX /diagnostics:column /sdl /Od /D _DEBUG /D _UNICODE /D UNICODE /Gm- /EHsc /MTd /GS /fp:precise /Zc:wchar_t /Zc:forScope /Zc:inline /permissive /external:W4 /Gd /TC /FC /J /utf-8 
$<$<CONFIG:RELEASE>:/Zi /nologo /W4 /WX /diagnostics:column /sdl /O2 /Oi /D NDEBUG /D _UNICODE /D UNICODE /Gm- /EHsc /MT /GS /Gy /fp:precise /Zc:wchar_t /Zc:forScope /Zc:inline /std:c++latest /permissive- /external:W4 /Gd /TP /FC /J /utf-8
)

target_link_options(xll PRIVATE /SUBSYSTEM:WINDOWS
$<$<CONFIG:DEBUG>:/NOLOGO /WX /MACHINE:X64 /SUBSYSTEM:WINDOWS
$<$<CONFIG:RELEASE>:/NOLOGO /WX /MACHINE:X64 /SUBSYSTEM:WINDOWS
)

