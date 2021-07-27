#include <iostream>
#include <QDir>
#include <QFileDialog>
#include <QTextCursor>
#include <QDebug>
#include <QStringListModel>

#include "libxl.h"

#include "BasicExcel.hpp"
//#include "NumberDuck.h"

#include "widget.h"
#include "ui_widget.h"
#include "util.h"


Widget::Widget(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::Widget)
{
    ui->setupUi(this);
}

Widget::~Widget()
{
    delete ui;
}

void Widget::checkInputItemColor()
{
    for (int i = 0; i < inputFilenameList.size(); ++i) {
        QString fileName = inputFilenameList[i];
        if (fileName.endsWith(QString(".xls"), Qt::CaseInsensitive)) {
            ui->listInputDirWidget->item(i)->setForeground(QColor("black"));
        } else {
            ui->listInputDirWidget->item(i)->setForeground(QColor("gray"));
        }
    }
}

void Widget::on_lsInputDirButton_clicked()
{
    //选择输入文件夹
    inputDirString = QDir::toNativeSeparators(QFileDialog::getExistingDirectory(this, "浏览文件夹", QDir::currentPath(), QFileDialog::DontUseNativeDialog));
    ui->inputDirLabel->setText(inputDirString);

    //获取输入文件夹下所有文件并展示出来
    QDir inputDir(inputDirString);
    inputDir.setFilter(QDir::Files);
    QFileInfoList inputFileList =  inputDir.entryInfoList();
    inputFilenameList.clear();
    actualInputFilenameList.clear();
    ui->listInputDirWidget->clear();
    for (int i = 0; i < inputFileList.size(); ++i) {
        QString fileName = inputFileList[i].fileName();
        ui->listInputDirWidget->addItem(fileName);
        inputFilenameList.append(fileName);
        if (fileName.endsWith(QString(".xls"), Qt::CaseInsensitive)) {
            actualInputFilenameList.append(fileName);
        }
    }
    checkInputItemColor();
    checkOutputItemColor();
}

void Widget::checkOutputItemColor()
{
    //如果输入文件夹中存在与输出文件夹中同名的文件且以".xls"结尾，则在输出文件夹中做特殊展示
    for (int i = 0; i < outputFilenameList.size(); ++i) {
        QString fileName = outputFilenameList[i];
        QList<QListWidgetItem*> inputItems = ui->listInputDirWidget->findItems(fileName, Qt::MatchFixedString);
        bool existSameNameEndsWithXls = false;
        for (QListWidgetItem* item : inputItems) {
            if (item->text().endsWith(".xls")) {
                existSameNameEndsWithXls = true;
            }
        }
        if (existSameNameEndsWithXls) {
            ui->listOutputDirWidget->item(i)->setForeground(QColor("red"));
        } else {
            ui->listOutputDirWidget->item(i)->setForeground(QColor("black"));
        }
    }
}

void Widget::on_lsOutputDirButton_clicked()
{
    //选择输出文件夹
    outputDirString = QDir::toNativeSeparators(QFileDialog::getExistingDirectory(this, "浏览文件夹", QDir::currentPath(), QFileDialog::DontUseNativeDialog));
    ui->outputDirLabel->setText(outputDirString);

    //获取输出文件夹下所有文件并展示出来
    QDir outputDir(outputDirString);
    outputDir.setFilter(QDir::Files);
    QFileInfoList outputFileList =  outputDir.entryInfoList();
    outputFilenameList.clear();
    ui->listOutputDirWidget->clear();
    for (int i = 0; i < outputFileList.size(); ++i) {
        QString fileName = outputFileList[i].fileName();
        ui->listOutputDirWidget->addItem(fileName);
        outputFilenameList.append(fileName);
    }
    checkOutputItemColor();
}

QString Widget::getOutputFile(const QString& inputFilename) {
    QDir outputDir(outputDirString);
    QString outputFile = outputDir.absoluteFilePath(inputFilename);
    return outputFile;
}


void Widget::on_dereplicateButton_clicked()
{
    QString outputFilenamePrefix = ui->outputFilenamePrefixLineEdit->text();
    QString outputFilenameSuffix = ui->outputFilenameSuffixLineEdit->text();
    //testXls4();


    QDir inputDir(inputDirString);
    YExcel::BasicExcel *srcExcel;
    YExcel::BasicExcel *dstExcel;
    for (int fileIndex = 0; fileIndex < actualInputFilenameList.size(); ++fileIndex) {
        QString inputFile = inputDir.absoluteFilePath(actualInputFilenameList[fileIndex]);
        QString outputFile = getOutputFile(actualInputFilenameList[fileIndex]);

        qDebug() << fileIndex << inputFile << outputFile;
        srcExcel = new YExcel::BasicExcel();
        dstExcel = new YExcel::BasicExcel();

        qDebug() << "Read Excel file from " << outputFile << " with address " << srcExcel;
        qDebug() << "dstExcel is initialized to " << dstExcel;
        readXlsFile(inputFile, srcExcel);
        int totalWorkSheets = srcExcel->GetTotalWorkSheets();

        vector<QString> tempsheetnames;

        for (int sheetIndex = 0; sheetIndex < totalWorkSheets; ++sheetIndex) {
            YExcel::BasicExcelWorksheet* srcSheet = srcExcel->GetWorksheet((size_t)sheetIndex);
            qDebug() << "Read file " << fileIndex << " and sheet " << sheetIndex << " with address " << srcSheet;
            vector<vector<QString>> uniqueValue = readUniqRows(srcSheet);

            YExcel::BasicExcelWorksheet* dstSheet;
            char* sheetNameChar = new char[1024];
            wchar_t* sheetNameWchar = new wchar_t[1024];
            if (srcExcel->GetSheetName(sheetIndex, sheetNameChar)) {
                dstSheet = dstExcel->AddWorksheet(sheetNameChar);
                dstSheet = dstExcel->GetWorksheet(sheetNameChar);
                qDebug() << "Get sheetNameChar as " << sheetNameChar << " with length " << strlen(sheetNameChar);
                tempsheetnames.push_back(QString(sheetNameChar));
            } else if (srcExcel->GetSheetName(sheetIndex, sheetNameWchar)){
                dstSheet = dstExcel->AddWorksheet(sheetNameWchar);
                dstSheet = dstExcel->GetWorksheet(sheetNameWchar);
                tempsheetnames.push_back(QString::fromWCharArray(sheetNameWchar));//这一步需要确认无问题
            } else {
                dstSheet = nullptr;
            }
            qDebug() << "Write sheet " << sheetIndex << " with address " << dstSheet;
            writeSheet(uniqueValue, dstSheet);

            // 这一步，成功写到dstSheet里了
            vector<vector<QString>> temp = readUniqRows(dstSheet);
            for (int i = 0; i < temp.size(); ++i) {
                for (int j = 0; j < temp[i].size(); ++j) {
                    //qDebug() << "TEST1: " << temp[i][j];
                }
            }
        }

        // 这一步，对整个表的每个sheet进行读取
        // 如果第二个文件是有3个sheet，且第2和第3个sheet内容为空，则发现生成的每个sheet都是0行0列
        // 如果第二个文件是有1个sheet，则发现生成的sheet大小为 2399183280  X  655384， 而且程序运行后出现异常
        // 这里重新从dstExcel获取的dstSheet地址，和之前读取表格写入dstExcel的dstSheet地址不同
        qDebug() << "TEST2: There are totally " << tempsheetnames.size() << " sheets to write.";
        for (int t = 0; t < tempsheetnames.size(); t++) {
            int dst_length = tempsheetnames[t].length() * 4;
            char* dst = new char[dst_length];
            //YExcel::BasicExcelWorksheet* dstSheet = dstExcel->GetWorksheet(Util::qstring2char(tempsheetnames[t], dst));
            //qDebug() << "TEST2: Read sheet " << t << " with name " << QString(dst) << " and address " << dstSheet;
            Util::qstring2char(tempsheetnames[t], dst);
            qDebug() << "Get sheetNameChar as " << dst << " with length " << strlen(dst);
            YExcel::BasicExcelWorksheet* dstSheet = dstExcel->GetWorksheet(t);
            qDebug() << "TEST2: Read sheet " << t << " with name " << tempsheetnames[t] << " and address " << dstSheet;
            vector<vector<QString>> temp;
            if (dstSheet) {
                temp = readUniqRows(dstSheet);
                qDebug() << "TEST2: sheet is read. ";
            } else {
                qDebug() << "TEST2: sheet is null.";
            }
            for (int i = 0; i < temp.size(); ++i) {
                for (int j = 0; j < temp[i].size(); ++j) {
                    //qDebug() << "TEST2: " << temp[i][j];
                }
            }
            delete dst;
        }

        qDebug() << "Write dstExcel to " << outputFile << " with address " << dstExcel;
        writeXlsFile(outputFile, dstExcel);
        delete srcExcel;
        delete dstExcel;
        qDebug() << "Done with file " << actualInputFilenameList[fileIndex] << ".";
    }

}

void Widget::testXls1()
{
    libxl::Book* book = xlCreateBook(); // xlCreateXMLBook() for xlsx
    if(book)
    {
        libxl::Sheet* sheet = book->addSheet(L"Sheet1");
        if(sheet)
        {
            sheet->writeStr(2, 1, L"Hello, World !");
            sheet->writeNum(3, 1, 1000);
        }
        book->save(L"example.xls");
        book->release();
    }
}

void Widget::testXls2()
{
    libxl::Book* book = xlCreateBook(); // xlCreateXMLBook() for xlsx
    std::wstring filename(L"D:\\QtProjects\\test.xls");
    qDebug() << "fuck 1";
    if(book->load(filename.c_str()))
    {
        qDebug() << "fuck 2";
        int sheetCount = book->sheetCount();
        for (int i = 0; i < sheetCount; ++i) {
            libxl::Sheet* sheet = book->getSheet(i);
            int firstRow = sheet->firstFilledRow();
            int lastRow = sheet->lastFilledRow();
            int firstColumn = sheet->firstFilledCol();
            int lastColumn = sheet->lastFilledCol();
            for (int j = firstRow; j <= lastRow; ++j) {
                for (int k = firstColumn; k <= lastColumn; ++k) {
                    qDebug() << "fuck " << j << " & " << k << ": " << sheet->readStr(j, k);
                }
            }
        }
        book->release();
    }
}

void Widget::readXlsFile(const QString& filename, YExcel::BasicExcel* excel) {
    char* dst = new char[filename.length() * 4];
    for (int i = 0; i < filename.length()*4; ++i) dst[i] = '\0';
    Util::qstring2char(QString(filename), dst);
    bool openSucceed = excel->Load(dst); //必须是 xls 类型的Excel文件，不能读取 xlsx 文件
    delete[] dst;
}

void Widget::writeXlsFile(const QString& filename, YExcel::BasicExcel* excel) {
    char* dst = new char[filename.length() * 4];
    for (int i = 0; i < filename.length()*4; ++i) dst[i] = '\0';
    Util::qstring2char(QString(filename), dst);
    bool writeSucceed = excel->SaveAs(dst); //必须是 xls 类型的Excel文件，不能写入 xlsx 文件
    delete[] dst;
}


vector<vector<QString>> Widget::readUniqRows(YExcel::BasicExcelWorksheet* sheet)
{
    vector<vector<QString>> result;
    vector<QString> resRow;
    if (sheet) {
        size_t maxRows = sheet->GetTotalRows(); //获取行数
        size_t maxCols = sheet->GetTotalCols(); //获取列数
        qDebug() << "Sheet size is " << maxRows << " X " << maxCols;
        for (int row = 0; row < maxRows; ++row) {
            resRow.clear();
            for (int col = 0; col < maxCols; ++col) {
                YExcel::BasicExcelCell* cell = sheet->Cell(row, col);
                switch (cell->Type())
                {
                  case YExcel::BasicExcelCell::UNDEFINED:
                    resRow.push_back(QString(""));
                    break;

                  case YExcel::BasicExcelCell::INT:
                    resRow.push_back(QString(cell->GetInteger()));
                    break;

                  case YExcel::BasicExcelCell::DOUBLE:
                    resRow.push_back(QString::number(cell->GetDouble(), 'f', 6));
                    break;

                  case YExcel::BasicExcelCell::STRING:
                    resRow.push_back(QString(cell->GetString()));
                    break;

                  case YExcel::BasicExcelCell::WSTRING:
                    //qDebug() << "fuck 3 " << row << " & " << col << ": " << QString::fromWCharArray(cell->GetWString());
                    resRow.push_back(QString::fromWCharArray(cell->GetWString()));
                    break;
                }
            }
            if (exist(resRow, result) == false) {
                result.push_back(resRow);
            }
        }
    }
    return result;
}

bool Widget::exist(const vector<QString>& resRow, const vector<vector<QString>>& source) {
    int maxRows = source.size();
    if (maxRows == 0) return false;
    int maxCols = maxRows ? source[0].size() : 0;
    for (int row = 0; row < maxRows; ++row) {
        if (source[row].size() == resRow.size()) {
            bool rowSame = true;
            for (int col = 0; col < maxCols; ++col) {
                rowSame = rowSame && (resRow[col].compare(source[row][col]) == 0);
                if (!rowSame) break;
            }
            if (rowSame) return true;
        }
    }
    return false;
}

bool Widget::writeSheet(const vector<vector<QString>>& source, YExcel::BasicExcelWorksheet* sheet)
{
    YExcel::BasicExcelCell* cell;
    int maxRows = source.size();
    int maxCols = maxRows ? source[0].size() : 0;
    if (sheet && maxRows && maxCols)
    {
        qDebug() << "Sheet size is " << maxRows << " X " << maxCols;
        for (int row = 0; row < maxRows; ++row) {
            for (int col = 0; col < maxCols; ++col) {
                int length = source[row][col].length() * 4;
                //char* valueChar = new char[length];
                //Util::qstring2char(source[row][col], valueChar);
                wchar_t* valueWchar = new wchar_t[length];
                for (int i = 0; i < length; ++i) valueWchar[i] = L'\0';
                source[row][col].toWCharArray(valueWchar);
                sheet->Cell(row, col)->SetWString(valueWchar);
                //qDebug() << "Write " << QString::fromWCharArray(valueWchar) << " to Cell[" << row << "][" << col <<"].";
            }
        }
        return true;
    } else {
        return false;
    }
}


void Widget::testXls3()
{
    YExcel::BasicExcel excel;

    //开始读取文件
    char* filename = "D:\\QtProjects\\示例\\0719.xls";
    QString filenameQstring(filename);
    char* dst = new char[filenameQstring.length() * 4];
    Util::qstring2char(QString(filename), dst);
    bool openSucceed = excel.Load(dst); //必须是 xls 类型的Excel文件，不能读取 xlsx 文件
    delete[] dst;

    //开始分析文件
    int totalWorkSheets = excel.GetTotalWorkSheets();
    //YExcel::BasicExcelWorksheet *sheet = excel.GetWorksheet("Sheet1"); //获取当前文件的指定名字的工作簿
    YExcel::BasicExcelWorksheet *sheet = excel.GetWorksheet((size_t)0); //获取当前文件的指定名字的工作簿
    if (sheet)
    {
        size_t maxRows = sheet->GetTotalRows(); //获取行数
        size_t maxCols = sheet->GetTotalCols(); //获取列数
        qDebug() << "行数=" << maxRows << ", 列数=" << maxCols;
        for (int i = 0; i < maxRows; ++i) {
            for (int j = 0; j < maxCols; ++j) {
                qDebug() << "fuck 3 " << i << " & " << j;
                YExcel::BasicExcelCell* cell = sheet->Cell(i,j);
                switch (cell->Type())
                {
                  case YExcel::BasicExcelCell::UNDEFINED:
                    qDebug() << "UNDEFINED";
                    printf(" ");
                    break;

                  case YExcel::BasicExcelCell::INT:
                    qDebug() << "INT" << cell->GetInteger();
                    printf("%10d", cell->GetInteger());
                    break;

                  case YExcel::BasicExcelCell::DOUBLE:
                    qDebug() << "DOUBLE" << cell->GetDouble();
                    printf("%10.6lf", cell->GetDouble());
                    break;

                  case YExcel::BasicExcelCell::STRING:
                    qDebug() << "STRING" << cell->GetString();
                    printf("%10s", cell->GetString());
                    break;

                  case YExcel::BasicExcelCell::WSTRING:
                    qDebug() << "WSTRING" << QString::fromWCharArray(cell->GetWString());
                    //wprintf(L"%10s", cell->GetWString());
                    break;
                }
            }
        }
    } else {
        qDebug() << "该文件似乎已被打开，请先关闭此文件。";
    }
    qDebug() << "结束";
}

void Widget::testListViewColor1()
{
    //
}


void Widget::testXls4()
{
    printf("Simple Example\n");
    printf("Create a spreadsheet!\n\n");

    /*
    NumberDuck::Workbook workbook("");
    NumberDuck::Worksheet* pWorksheet = workbook.GetWorksheetByIndex(0);

    NumberDuck::Cell* pCell = pWorksheet->GetCellByAddress("A1");
    pCell->SetString("Totally cool spreadsheet!");

    pWorksheet->GetCell(1,1)->SetFloat(3.1417f);

    workbook.Save("SimpleExample.xls");

    NumberDuck::Workbook* pWorkbookIn = new NumberDuck::Workbook("");
    if (pWorkbookIn->Load("SimpleExample.xls"))
    {
        NumberDuck::Worksheet* pWorksheetIn = pWorkbookIn->GetWorksheetByIndex(0);
        NumberDuck::Cell* pCellIn = pWorksheetIn->GetCell(0,0);
        printf("Cell Contents: %s\n", pCellIn->GetString());
    }
    */
}







