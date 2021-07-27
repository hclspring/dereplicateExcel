#ifndef WIDGET_H
#define WIDGET_H

#include <QWidget>

#include "BasicExcel.hpp"
//#include "NumberDuck.h"

QT_BEGIN_NAMESPACE
namespace Ui { class Widget; }
QT_END_NAMESPACE

class Widget : public QWidget
{
    Q_OBJECT

public:
    Widget(QWidget *parent = nullptr);
    ~Widget();

    void testXls1();
    void testXls2();
    void testXls3();
    void testXls4();
    void testListViewColor1();

private slots:
    void on_lsInputDirButton_clicked();

    void on_lsOutputDirButton_clicked();

    void on_dereplicateButton_clicked();

private:
    Ui::Widget *ui;
    QString inputDirString;
    QString outputDirString;
    //QString outputFilenamePrefix;
    //QString outputFilenameSuffix;
    QStringList inputFilenameList;
    QStringList outputFilenameList;
    QStringList actualInputFilenameList;
    //QStringList actualOutputFilenameList;

private:
    void checkInputItemColor();
    void checkOutputItemColor();
    QString getOutputFile(const QString& inputFilename);

    void readXlsFile(const QString& filename, YExcel::BasicExcel* excel);
    void writeXlsFile(const QString& filename, YExcel::BasicExcel* excel);
    vector<vector<QString>> readUniqRows(YExcel::BasicExcelWorksheet* sheet);
    bool exist(const vector<QString>& resRow, const vector<vector<QString>>& source);
    bool writeSheet(const vector<vector<QString>>& source, YExcel::BasicExcelWorksheet* sheet);
};
#endif // WIDGET_H
