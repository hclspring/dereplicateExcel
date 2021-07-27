#include "widget.h"

#include <locale>
#include <QApplication>
#include <QTextCodec>
#include <QDebug>

//#if defined(_MSC_VER) && (_MSC_VER >= 1600)
//# pragma execution_character_set("utf-8")
//#endif

int main(int argc, char *argv[])
{
    //使用setlocale函数将本机的语言设置为中文简体
    //LC_ALL表示设置所有的选项（包括金融货币、小数点，时间日期格式、语言字符串的使用习惯等），chs表示中文简体
    //setlocale(LC_ALL, "chs");
    //setlocale(LC_ALL,"Chinese-simplified");
    QApplication a(argc, argv);

    //设置中文字体
    a.setFont(QFont("Microsoft Yahei", 9));

//设置中文编码
#if (QT_VERSION <= QT_VERSION_CHECK(5,0,0))
#if _MSC_VER
    QTextCodec *codec = QTextCodec::codecForName("gbk");
#else
    QTextCodec *codec = QTextCodec::codecForName("utf-8");
#endif
    QTextCodec::setCodecForLocale(codec);
    QTextCodec::setCodecForCStrings(codec);
    QTextCodec::setCodecForTr(codec);
#else
    QTextCodec *codec = QTextCodec::codecForName("utf-8");
    QTextCodec::setCodecForLocale(codec);
#endif

//    QTextCodec *codec = QTextCodec::codecForName("UTF-8");
//    QTextCodec::setCodecForLocale(codec);
//    QTextCodec::setCodecForCStrings(codec);
//    QTextCodec::setCodecForTr(codec);

    Widget w;
    w.show();
    return a.exec();
}
