cmake_minimum_required(VERSION 3.20)

project(EmbedPack LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

add_executable(EmbedPack WIN32
    main.cpp
    App.cpp
    CoreServices.cpp
)

target_compile_definitions(EmbedPack PRIVATE
    UNICODE
    _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
)

if (WIN32)
    target_compile_definitions(EmbedPack PRIVATE
        _WIN32_WINNT=0x0A00
        WINVER=0x0A00
        NTDDI_VERSION=0x0A000000
    )
endif()

if (MSVC)
    target_compile_options(EmbedPack PRIVATE
        /W4
        /permissive-
        /Zc:__cplusplus
        /Zc:wchar_t
        /EHsc
        /utf-8
        /MP
    )

    set_property(TARGET EmbedPack PROPERTY
        MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>"
    )
else()
    target_compile_options(EmbedPack PRIVATE
        -Wall -Wextra -Wpedantic
        -Wconversion -Wsign-conversion
    )
endif()

target_link_libraries(EmbedPack PRIVATE
    comdlg32
    user32
    gdi32
    kernel32
    comctl32
)
