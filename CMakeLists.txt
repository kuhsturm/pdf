cmake_minimum_required(VERSION 3.16)
project(ElektrikRaporSistemi VERSION 1.0.0 LANGUAGES CXX)

# C++17 standardı
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

# Qt6 bul
find_package(Qt6 REQUIRED COMPONENTS Core Gui Widgets Pdf Xml)

# QXlsx - subdirectory olarak dahil et
set(QXLSX_PARENTPATH ${CMAKE_SOURCE_DIR}/libs/QXlsx)
add_subdirectory(${CMAKE_SOURCE_DIR}/libs/QXlsx/QXlsx QXlsx)

# Kaynak dosyaları
set(SOURCES
    src/main.cpp
    # Core
    src/core/ExcelReader.cpp
    src/core/DocxWriter.cpp
    src/core/PdfParser.cpp
    src/core/FlukeExtractor.cpp
    src/core/KisiBilgileriReader.cpp
    # Logic
    src/logic/ReportGenerator.cpp
    src/logic/Validators.cpp
    # GUI
    src/gui/MainWindow.cpp
    src/gui/PanoTabWidget.cpp
    src/gui/FonksiyonTestleriTable.cpp
    src/gui/DragDropWidget.cpp
    src/gui/CommonInfoPanel.cpp
    src/gui/GozleKontrolWidget.cpp
    # Dialogs
    src/gui/dialogs/LinieGroupDialog.cpp
    src/gui/dialogs/SettingsDialog.cpp
    resources.qrc
)

set(HEADERS
    src/core/ExcelReader.h
    src/core/DocxWriter.h
    src/core/PdfParser.h
    src/core/FlukeExtractor.h
    src/core/KisiBilgileriReader.h
    src/logic/DataModels.h
    src/logic/ReportGenerator.h
    src/logic/Validators.h
    src/gui/MainWindow.h
    src/gui/PanoTabWidget.h
    src/gui/FonksiyonTestleriTable.h
    src/gui/DragDropWidget.h
    src/gui/CommonInfoPanel.h
    src/gui/GozleKontrolWidget.h
    src/gui/dialogs/LinieGroupDialog.h
    src/gui/dialogs/SettingsDialog.h
)

# Executable oluştur
add_executable(${PROJECT_NAME} WIN32
    ${SOURCES}
    ${HEADERS}
)

# Include dizinleri
target_include_directories(${PROJECT_NAME} PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${CMAKE_SOURCE_DIR}/src/core
    ${CMAKE_SOURCE_DIR}/src/logic
    ${CMAKE_SOURCE_DIR}/src/gui
)

# Qt ve QXlsx ile linkle
target_link_libraries(${PROJECT_NAME} PRIVATE
    Qt6::Core
    Qt6::Gui
    Qt6::Widgets
    Qt6::Pdf
    Qt6::Xml
    QXlsx
)

# QXlsx header'larını include et
target_include_directories(${PROJECT_NAME} PRIVATE
    ${CMAKE_SOURCE_DIR}/libs/QXlsx/QXlsx/header
)

# Windows için: Konsol gizle, icon ekle
if(WIN32)
    set_target_properties(${PROJECT_NAME} PROPERTIES
        WIN32_EXECUTABLE TRUE
    )
    # Static CRT linkage (tek EXE için)
    set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
endif()

# Release optimizasyonları
if(CMAKE_BUILD_TYPE STREQUAL "Release")
    if(MSVC)
        target_compile_options(${PROJECT_NAME} PRIVATE /O2 /GL)
        target_link_options(${PROJECT_NAME} PRIVATE /LTCG)
    else()
        target_compile_options(${PROJECT_NAME} PRIVATE -O3)
    endif()
endif()

# Install kuralları
install(TARGETS ${PROJECT_NAME} RUNTIME DESTINATION bin)
install(DIRECTORY ${CMAKE_SOURCE_DIR}/sablon DESTINATION bin)
