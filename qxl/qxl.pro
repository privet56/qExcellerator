QT       += core gui

include(../QtXlsxWriter-master/src/xlsx/qtxlsx.pri)
#QT+= xlsx

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

TARGET = qexcellerator
TEMPLATE = app

SOURCES += main.cpp\
        mainwindow.cpp \
    xlsxsheetmodel.cpp \
    sheet.cpp

HEADERS  += mainwindow.h \
    xlsxsheetmodel.h \
    xlsxsheetmodel_p.h \
    res.rc \
    resource.h \
    sheet.h

FORMS    += mainwindow.ui

CONFIG += mobility
MOBILITY = 

RESOURCES += res.qrc

INCLUDEPATH += ../build-qtxlsx-5_5_desktop_64bit_vs2013-Release/include/QtXlsx
debug:LIBS += ../build-qtxlsx-5_5_desktop_64bit_vs2013-Debug/lib/Qt5Xlsx.lib
release:LIBS += ../build-qtxlsx-5_5_desktop_64bit_vs2013-Release/lib/Qt5Xlsx.lib

DISTFILES += ico.ico

RC_FILE = res.rc
