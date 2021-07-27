QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++11

#QMAKE_CXXFLAGS += \\utf-8

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    BasicExcel.cpp \
    main.cpp \
    widget.cpp

HEADERS += \
    BasicExcel.hpp \
    excel.h \
    util.h \
    widget.h

FORMS += \
    widget.ui

INCLUDEPATH += D:\\Qt\\libxl-3.9.4.3\\include_cpp
LIBS += D:\\Qt\\libxl-3.9.4.3\\lib\\libxl.lib

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target



#win32:CONFIG(release, debug|release): LIBS += -L$$PWD/../../Qt/NumberDuck_v2.4.4/Bin32/ -lNumberDuck
#else:win32:CONFIG(debug, debug|release): LIBS += -L$$PWD/../../Qt/NumberDuck_v2.4.4/Bin32/ -lNumberDuckd
#else:unix: LIBS += -L$$PWD/../../Qt/NumberDuck_v2.4.4/Bin32/ -lNumberDuck
#
#INCLUDEPATH += $$PWD/../../Qt/NumberDuck_v2.4.4/Include
#DEPENDPATH += $$PWD/../../Qt/NumberDuck_v2.4.4/Bin32
