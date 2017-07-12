/* This program is not able to consistently output the save file when saving directly to
 * a server, so make sure to save everything locally and move it to the correct location
 * afterwards manually.  Making a copy of all of the files used for backup may be helpful.  */

#include "mainwindow.h"
#include "ui_mainwindow.h"

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
    ui->lblStatus->setText("Opening file...");
    ui->lblStatus->repaint();
    readFromXlsx(ui->fileUrl->text());
}

void MainWindow::on_btnReadOutput_clicked()
{ // Reads from the file location determined by the template file and the save name selected
    QString fullOutputUrl = ui->fileUrl->text().split('/').back();
    fullOutputUrl = ui->fileUrl->text().remove(
                ui->fileUrl->text().length()-fullOutputUrl.length(), fullOutputUrl.length());
    if(ui->outputUrl->text().compare("") == 0)
        fullOutputUrl += ui->lineId->text() + ".xlsx";
    else
        fullOutputUrl += ui->outputUrl->text() + ".xlsx";

    ui->lblStatus->setText("Opening file...");
    ui->lblStatus->repaint();
    readFromXlsx(fullOutputUrl);
}

void MainWindow::readFromXlsx(QString fileUrl)
{ // Opens the excel file at location fileUrl for reading or editing.
    QAxObject excel("Excel.Application");
    QFile f(fileUrl);
    if(excel.isNull() || !f.exists())
    { // If the file does not exist, don't read the file
        ui->lblStatus->setText("Failed to read file.");
        return;
    }
    f.deleteLater();

    excel.setProperty("Visible", true);

    QAxObject *workbooks = excel.querySubObject("WorkBooks");
    workbooks->dynamicCall("Open (const QString&)", fileUrl);
    ui->lblStatus->setText("Successfully read file.");
}

void MainWindow::on_btnWrite_clicked()
{
    // First, check if any of the necessary lines are empty
    if(ui->txtUrl->text().compare("") == 0 ||
            ui->fileUrl->text().compare("") == 0 ||
            ui->lineId->text().compare("") == 0)
    {
        ui->lblStatus->setText("Error: must fill in all fields marked with a * .");
        return;
    }
    QVariant outputUrl = ui->outputUrl->text();
    if(outputUrl.compare("") == 0)
    { // If the "Save name" line is empty, set the output URL to the ID name
        outputUrl = ui->lineId->text();
        ui->outputUrl->setPlaceholderText(outputUrl.toString());
    }

    ui->lblStatus->setText("Locating text file...");
    ui->lblStatus->repaint();

    QFile *txtFile = new QFile(ui->txtUrl->text());
    if(!txtFile->exists())
    { // Abort if the data file does not exist
        ui->lblStatus->setText("Error: Failed to read file.");
        return;
    }

    txtFile->open(QIODevice::ReadOnly);
    if(!txtFile->peek(1).startsWith("="))
    { /* Reformat the text file so that it can be read properly unless user doesn't want to
         overwrite an existing reformatted text file if such file exists */
        if(!reformatTxt(txtFile))
        {
            ui->lblStatus->setText("Aborted text file reformatting.  Canceled write.");
            return;
        }
    }
    txtFile->close();

    ui->lblStatus->setText("Writing excel file...");
    ui->lblStatus->repaint();

    // Set fulloutputUrl to be the full directory that the output file will be saved to
    QString fullOutputUrl = ui->fileUrl->text().split('/').back();
    fullOutputUrl = ui->fileUrl->text().remove(
                ui->fileUrl->text().length()-fullOutputUrl.length(), fullOutputUrl.length());
    fullOutputUrl += outputUrl.toString() + ".xlsx";

    // Execute the writeToXlsx method and store the status message in s.
    QString s = writeToXlsx(txtFile, ui->lineId->text(),
                   ui->fileUrl->text(), fullOutputUrl);
    ui->lblStatus->setText(s);
}

bool MainWindow::reformatTxt(QFile *txtFile)
{ // Reformats the .txt file and saves it to a separate file in the same directory
    ui->lblStatus->setText("Reformatting text file...");
    ui->lblStatus->repaint();

    // Make the reformatted file a separate file in the same folder
    QString newName = ui->txtUrl->text().split('/').back();
    newName = ui->txtUrl->text().remove(
                ui->txtUrl->text().length()-4, 4);
    newName += " reformatted.txt";

    if(QFile(newName.toLocal8Bit()).exists())
    { // If the file has already been reformatted, ask if the user wants to overwrite it
        QMessageBox confirmBox;
        confirmBox.setText("The reformatted text file for the data already exists.");
        confirmBox.setInformativeText("Would you like to continue and overwrite this file?  \
If you plan on using this reformatted text file repeatedly, consider changing the \
text file directory to the _reformatted.txt.");
        confirmBox.setWindowTitle("Confirm Overwrite");
        confirmBox.setStandardButtons(QMessageBox::Yes | QMessageBox::No);
        confirmBox.setDefaultButton(QMessageBox::No);
        if(confirmBox.exec() == QMessageBox::No)
            return false;
    }

    QByteArray newBA;
    while(!txtFile->atEnd())
        newBA.append(txtFile->readLine().split(' ').back());
    txtFile->close();

    txtFile->setFileName(newName.toLocal8Bit());

    // Append the desired text into newBA to be written to the new file
    txtFile->open(QIODevice::WriteOnly);
    txtFile->write(newBA);
    txtFile->close();
    return true;
}

QString MainWindow::writeToXlsx(QFile *txtFile, QString id, QString excelUrl,  QString outputUrl)
{   /* Finds the URL of the excel file to write the contents of the text file to with the
    correct id and saves the edited file to the output URL */

    txtFile->open(QIODevice::ReadOnly);

    if(!txtFile->readLine().startsWith("=")) // Text file should start with "="
        return tr("Error: File doesn't start with '='.");
    if(!findNextColumn(txtFile, id)) // Text file should have at least one matching entry
        return tr("Error: File does not have any entries with the matching ID.");

    QAxObject excel("Excel.Application");
    if(excel.isNull()) // Excel file should exist
        return tr("Error: The excel file is null.");

    //excel.setProperty("Visible", true);

    QAxObject *workbooks = excel.querySubObject("WorkBooks");
        workbooks->dynamicCall("Open (const QString&)", excelUrl);
    QAxObject *workbook = excel.querySubObject("ActiveWorkBook");

    // Gets data from the 1st worksheet (worksheets are numbered starting from 1)
    QAxObject *worksheet = workbook->querySubObject("Worksheets(int)", 1);

    QAxObject *range;
    QString cell;
    QString cellStr;
    int col = 0;
    int row = 0;

    while(txtFile->readLine().startsWith("start"))
    { // While there are more columns with the matching ID
        cell = "";
        row = 1;
        cell.append(numToAlph(col).append(QString::fromStdString(std::to_string(row))));
        range = worksheet->querySubObject("Range(const QString &)", cell);
        range->setProperty("Value", "start");
        while(!txtFile->peek(1).startsWith("=") && !txtFile->atEnd())
        { // While haven't run into "=" or eof
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

    // Set the ID of the graph to the appropriate name
    cell = numToAlph(col+2).append("2");
    range = worksheet->querySubObject("Range(const QString &)", cell);
    range->setProperty("Value", id);

/*    QAxObject *columns = worksheet->querySubObject("Columns(cons QString&:const QString&)",
 *                                                    "A", numToAlph(col));
 *    columns->dynamicCall("AutoFit()"); // Auto-fits the text so it is shown fully */

    // Save and close the excel file
    //excel.setProperty ("DisplayAlerts", 0);
    workbook->dynamicCall("SaveCopyAs (const QString&)", outputUrl);
    //excel.setProperty("DisplayAlerts", 1);
    workbook->dynamicCall("Close (Boolean)", false);
    excel.dynamicCall("Quit (void)");
    return "Successfully written.";
}

bool MainWindow::findNextColumn(QFile *txtFile, QString id)
{ // Useful for when different entries with the same ID in the text file are not ordered
    QString nextLine;
    while(true)
    { // Finds the next entry in txtFile with the correct ID
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
{   // 0 is A, 26 is AA.  These are the column names in excel documents.
    QString output = "";
    QChar curNum;
    while(num >= 26)
    {
        curNum = QChar((num % 26)+65);
        output += curNum;
        num /= 26;
    }
    output += QChar(num+65); // ASCII for A = 65
    return output;
}

void MainWindow::on_btnTxtFile_clicked()
{
    QFileDialog txtDialog;
    ui->txtUrl->setText(txtDialog.getOpenFileName(this, "Find your Text file", NULL,
                                                   "Text Files (*.txt);;All (*.*)"));
}

void MainWindow::on_btnExcelFile_clicked()
{
    QFileDialog xlDialog;
    ui->fileUrl->setText(xlDialog.getOpenFileName(this, "Find your Excel file", NULL,
                                                   "Excel (*.xls *.xlsx);;All (*.*)"));
}

void MainWindow::on_outputUrl_editingFinished()
{
    QString newTxt = ui->outputUrl->text();
    if(newTxt.compare("") == 0)
    {
        newTxt = "Write File as " + ui->lineId->text() + ".xlsx";
        ui->outputUrl->setPlaceholderText(ui->lineId->text());
    }
    else
        newTxt = "Write File as " + newTxt + ".xlsx";
    ui->btnWrite->setText(newTxt);
}

void MainWindow::on_lineId_editingFinished()
{
    if(ui->outputUrl->text().compare("") == 0)
    {
        ui->outputUrl->setPlaceholderText(ui->lineId->text());
        ui->btnWrite->setText("Write File as " + ui->lineId->text() + ".xlsx");
    }
}

void MainWindow::on_actionAbout_triggered()
{
    QMessageBox aboutBox;
    aboutBox.setWindowTitle("About");
    aboutBox.setText("How to use:");
    aboutBox.setInformativeText("Select a text file with the data that you want to insert into \
the excel template file.\nClicking on the above buttons opens a file browser to make finding each \
of the files easier.\nThen, enter the ID of the data in the text file that you want to extract.\n\
Also enter a file name that the edited excel will be saved to.\nAll fields marked with a * must \
be filled in.\nIf the save name is not filled in, it will default to the name of the ID.");
    aboutBox.exec();
}
