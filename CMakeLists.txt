cmake_minimum_required(VERSION 3.12)
project(xll
	VERSION 0.1.0
	DESCRIPTION "A library for creating Excel add-ins."
	HOMEPAGE_URL "https://github.com/xlladdins/xll24"
)

set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED True)

add_library(STATIC
addin.cpp
debug.cpp
dllmain.cpp
evaluate.cpp
fpx.c
paste.cpp
range.cpp
xlauto.cpp
)
target_compile_options(xll PRIVATE 
$<$<CONFIG:DEBUG>:-g>
$<$<CONFIG:RELEASE>:-O3>
)
target_link_options(xll PRIVATE /SUBSYSTEM:WINDOWS
$<$<CONFIG:DEBUG>:-g>
$<$<CONFIG:RELEASE>:-O3>
)
