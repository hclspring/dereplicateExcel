#ifndef UTIL_H
#define UTIL_H

#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <locale>
#include <QString>
#include <QDebug>

/*
 * Usage:
 * 1.wchar_t* to QString:
 *          QString::fromWCharArray(wchar_t*)
 *
 * 2.wchar_t* to char*:
 *          dst_length = source_length * 4;
 *          char* dst = new char[dst_length];
 *          bool result = wchar2char(wchar_t* source, char* dst, int dst_length);
 *
 * 3.char* to wchar_t*:
 *
 * 4.char* to QString:
 *          QString dst = QString(char* source);
 *
 * 5.QString to char*:
 *          dst_length = source.length() * 4;
 *          char* dst = new char[dst_length];
 *          bool result = qstring2char(const QString& source, char* dst);
 *
 * 6.QString to wchar*:
 *
 */

class Util {
public:
    static bool wchar2char(const wchar_t* source, char* dst, int dst_length)
    // PURPOSE: transfer wchar_t* string to char* string, with Chinese.
    // ATTENTION: memory for dst is applicated before this function, it should be freed afterwards.
    // PROMISE: Return true if source is successfully transferred, false if otherwise.
    {
        for (int i = 0; i < dst_length; ++i) dst[i] = '\0';
        setlocale(LC_ALL, "Chinese-simplified");
        size_t temp_size;
        wcstombs_s(&temp_size, dst, dst_length, source, _TRUNCATE);
        return true;
    }

    static bool qstring2char(const QString& source, char* dst)
    // PURPOSE: transfer QString to char* string, with Chinese.
    // ATTENTION: memory for dst is applicated *IN* this function, it should be freed afterwards.
    //          That is, when dst is unuseful, you should run "delete[] dst;"
    {
        wchar_t* wa = new wchar_t[source.length()*4];
        for (int i = 0; i < source.length()*4; ++i) wa[i] = L'\0';
        int wl = source.toWCharArray(wa);
        //std::wcout << "qstring2char: " << wa << std::endl;
        int cl = wl * 4;
        //dst = new char[cl];
        wchar2char(wa, dst, cl);
        delete[] wa;
    }

    /*
    // 此函数没啥用
    static bool wchar2qstring(const wchar_t* source, QString& qstring)
    {
        int dst_length = wcslen(source) * 4;
        char* dst = new char[dst_length];
        wchar2char(source, dst, dst_length);
        qstring = QString(dst);
        delete[] dst;
        return true;
    }
    */
};

#endif // UTIL_H
