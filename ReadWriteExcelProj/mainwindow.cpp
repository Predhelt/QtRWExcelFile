/* This program is not able to consistently output the save file directly to a server,
 * so make sure to save everything locally and move it to the correct location afterwards
 * manually.  Making a copy of all of the files used for backup may be helpful.  */

#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
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

void MainWindow::on_btnReadTemplate_clicked()
{
    if(readFromXlsx(ui->fileUrl->text()))
        ui->lblStatus->setText("Successfully read file.");
    else
        ui->lblStatus->setText("Failed to read file.");
}

void MainWindow::on_btnReadOutput_clicked()
{
    QString fullOutputUrl = ui->fileUrl->text().split('/').back();
    fullOutputUrl = ui->fileUrl->text().remove(
                ui->fileUrl->text().length()-fullOutputUrl.length(), fullOutputUrl.length());
    fullOutputUrl += ui->outputUrl->text() + ".xlsx";

    if(readFromXlsx(fullOutputUrl))
        ui->lblStatus->setText("Successfully read file.");
    else
        ui->lblStatus->setText("Failed to read file.");
}

bool MainWindow::readFromXlsx(QString fileUrl)
{
    QAxObject excel("Excel.Application");
    QFile f(fileUrl);
    if(excel.isNull() || !f.exists())
    {
        return false;
    }
    f.deleteLater();

    excel.setProperty("Visible", true);

    QAxObject *workbooks = excel.querySubObject("WorkBooks");
    workbooks->dynamicCall("Open (const QString&)", fileUrl);
    return true;
}

void MainWindow::on_btnWrite_clicked()
{
    ui->lblStatus->setText("Locating text file...");
    ui->lblStatus->repaint();

    QFile *txtFile = new QFile(ui->txtUrl->text());
    if(!txtFile->exists())
    {
        ui->lblStatus->setText("Error: Failed to read file.");
        return;
    }

    txtFile->open(QIODevice::ReadOnly);
    if(!txtFile->peek(1).startsWith("="))
    { //Reformat the text file so that it can be read properly
        ui->lblStatus->setText("Reformatting text file...");
        ui->lblStatus->repaint();

        QByteArray newBA;
        while(!txtFile->atEnd())
        {
            newBA.append(txtFile->readLine().split(' ').back());
        }
        txtFile->close();

        //Make the reformatted file a separate file in the same folder
        QString newName = ui->txtUrl->text().split('/').back();
        newName = ui->txtUrl->text().remove(
                    ui->txtUrl->text().length()-4, 4);
        newName += " reformatted.txt";

        txtFile->setFileName(newName.toLocal8Bit());
        txtFile->open(QIODevice::WriteOnly);
        txtFile->write(newBA);
        txtFile->close();
    }

    ui->lblStatus->setText("Writing excel file...");
    ui->lblStatus->repaint();

    QString fullOutputUrl = ui->fileUrl->text().split('/').back();
    fullOutputUrl = ui->fileUrl->text().remove(
                ui->fileUrl->text().length()-fullOutputUrl.length(), fullOutputUrl.length());
    fullOutputUrl += ui->outputUrl->text() + ".xlsx";

    if(writeToXlsx(txtFile, ui->lineId->text(),
                   ui->fileUrl->text(), fullOutputUrl))
        ui->lblStatus->setText("Successfully written.");
    else
        ui->lblStatus->setText("Error: Failed to write to file.");
}

bool MainWindow::writeToXlsx(QFile *txtFile, QString id, QString excelUrl,  QString outputUrl)
{   /* finds the url of the excel file to write the contents of the text file to with the
    correct id and saves the edited file to the output url */

    txtFile->open(QIODevice::ReadOnly);

    if(!txtFile->readLine().startsWith("=")) //text file should start with "="
    {
        qDebug() << tr("Error: File doesn't start with '='.");
        return false;
    }
    if(!findNextColumn(txtFile, id)) //Text file should have at least one matching entry
    {
        qDebug() << tr("Error: File does not have any entries with the matching id.");
    }

    QAxObject excel("Excel.Application");
    if(excel.isNull()) //Excel file should exist
    {
        qDebug() << tr("Error: The excel file is null.");
        return false;
    }

    //excel.setProperty("Visible", true);

    QAxObject *workbooks = excel.querySubObject("WorkBooks");
        workbooks->dynamicCall("Open (const QString&)", excelUrl);
    QAxObject *workbook = excel.querySubObject("ActiveWorkBook");

    //gets data from the 1st worksheet (worksheets are numbered starting from 1)
    QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);

    QAxObject *range;
    QString cell;
    QString cellStr;
    int col = 0;
    int row = 0;

    while(txtFile->readLine().startsWith("start"))
    { //while there are more columns with the matching id
        cell = "";
        row = 1;
        cell.append(numToAlph(col).append(QString::fromStdString(std::to_string(row))));
        range = worksheet->querySubObject("Range(const QString &)", cell);
        range->setProperty("Value", "start");
        while(!txtFile->peek(1).startsWith("=") && !txtFile->atEnd())
        { //while haven't run into "=" or eof
            cell = "";
            cell.append(numToAlph(col).append(QString::fromStdString(std::to_string(row+1))));
            range = worksheet->querySubObject("Range(const QString &)", cell);

            cellStr = txtFile->readLine().split('\r').front();
            range->setProperty("Value", cellStr);
            row++;
        }
        txtFile->readLine();
        if(!findNextColumn(txtFile, id))
            break;
        col++;
    }

    //Set the ID of the graph to the appropriate name
    cell = numToAlph(col+2).append("2");
    range = worksheet->querySubObject("Range(const QString &)", cell);
    range->setProperty("Value", id);

/*    QAxObject *columns = worksheet->querySubObject("Columns(cons QString&:const QString&)",
 *                                                    "A", numToAlph(col));
 *    columns->dynamicCall("AutoFit()"); */

    // Save and close the excel file
    //excel.setProperty ("DisplayAlerts", 0);
    workbook->dynamicCall("SaveCopyAs (const QString&)", outputUrl);
    //excel.setProperty("DisplayAlerts", 1);
    workbook->dynamicCall("Close (Boolean)", false);
    excel.dynamicCall("Quit (void)");
    return true;
}

bool MainWindow::findNextColumn(QFile *txtFile, QString id)
{
    QString nextLine;
    while(true)
    { //Finds the next entry in txtFile with the correct ID
        nextLine = QString(txtFile->readLine().split('\r').front());
        if(!id.compare(nextLine))
        {
            txtFile->readLine(); txtFile->readLine(); break;
        }
        else
        {
            while(!txtFile->readLine().startsWith("="))
                if(txtFile->atEnd())
                    return false;
        }
    }
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

void MainWindow::on_btnTxtFile_clicked()
{
    QFileDialog txtDialog;
    ui->txtUrl->setText(txtDialog.getOpenFileName(this, "Find your Text file", "",
                                                   "Text Files (*.txt);;All (*.*)"));
}

void MainWindow::on_btnExcelFile_clicked()
{
    QFileDialog txtDialog;
    ui->fileUrl->setText(txtDialog.getOpenFileName(this, "Find your Excel file", "",
                                                   "Excel (*.xls *.xlsx);;All (*.*)"));
}

void MainWindow::on_outputUrl_editingFinished()
{
    QString newTxt = "Write File as " + ui->outputUrl->text() + ".xlsx";
    ui->btnWrite->setText(newTxt);
}

void MainWindow::on_actionAbout_triggered()
{
    ui->lblStatus->setText("To use: select a text file with the data that you want to insert \
into the excel template file.  Clicking on the buttons opens a file browser to make finding \
the files easier.  Then, enter the id of the data in the text file that you want to extract.  \
  Also enter a file name that the edited excel will be saved to.  These cannot be left blank.");
}
