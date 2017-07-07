#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QTextStream>
#include <QAxobject>
#include <QException>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_btnRead_clicked()
{
    readFromXlsx(ui->outputUrl->text());
}

void MainWindow::readFromXlsx(QString fileUrl)
{
    QAxObject excel("Excel.Application");
    if(excel.isNull())
        return;

    excel.setProperty("Visible", true);

    QAxObject *workbooks = excel.querySubObject("WorkBooks");
    workbooks->dynamicCall("Open (const QString&)", fileUrl);
}

void MainWindow::on_btnWrite_clicked()
{
    ui->lblStatus->setText("Writing...");
    ui->lblStatus->repaint();

    if(writeToXlsx(ui->fileUrl->text(), ui->outputUrl->text()))
        ui->lblStatus->setText("Successfully written");
    else
        ui->lblStatus->setText("Failed to write table");
}



bool MainWindow::writeToXlsx(QString fileUrl, QString outputUrl)
{
    QAxObject excel("Excel.Application");
    if(excel.isNull())
        return false;

    excel.setProperty("Visible", true);

    QAxObject *workbooks = excel.querySubObject("WorkBooks");
        workbooks->dynamicCall("Open (const QString&)", fileUrl);
    QAxObject *workbook = excel.querySubObject("ActiveWorkBook");

    // I need data from the 1st worksheet (worksheets are numbered starting from 1)
    QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);

    QAxObject *range;
    QString cell;
    QString tblCell;
    QTableWidget *table = ui->tableWidget;
    for(int row = 0; row < table->rowCount(); row++)
    {
        for(int col = 0; col < table->columnCount(); col++)
        {
            cell = "";
            cell.append(numToAlph(col).append(QString::fromStdString(std::to_string(row+1))));
            range = worksheet->querySubObject("Range(const QString &)", cell);
            if(table->item(row,col))
                tblCell = table->item(row, col)->text();
            else
                tblCell = "";
            range->setProperty("Value", tblCell);
        }
    }
    QAxObject *columns = worksheet->querySubObject("Columns(cons QString&:const QString&)",
                                                   "A", numToAlph(table->columnCount()-1));
    columns->dynamicCall("AutoFit()");

    // Save and close the excel file
    //excel.setProperty ("DisplayAlerts", 0);
    workbook->dynamicCall("SaveCopyAs (const QString&)", outputUrl);
    //excel.setProperty("DisplayAlerts", 1);
    workbook->dynamicCall("Close (Boolean)", false);
    excel.dynamicCall("Quit (void)");
    return true;
}

QString MainWindow::numToAlph(int num)
{   //0 is A, 26 is AA
    QString output = "";
    QChar curNum;
    while(num >= 26)
    {
        curNum = QChar((num % 26)+65);
        output += curNum;
        num /= 26;
    }
    output += QChar(num+65); //ASCII for A = 65
    return output;
}
