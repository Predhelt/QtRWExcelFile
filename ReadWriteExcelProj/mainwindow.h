/* This project asks for the user to input the location of an existing text
 * file with data to be added to an existing excel file template whose location
 * must also be given.  The data with the given ID in the text file will be
 * written into the template file and saved to a new excel file in the same
 * directory as the template file with either the given save name or the ID
 * of the extracted data if no save name is given. */

#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFile>
#include <QDebug>
#include <QFileDialog>
#include <QTextStream>
#include <QAxobject>
#include <QException>
#include <QMessageBox>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void on_btnReadTemplate_clicked();
    void on_btnReadOutput_clicked();

    void on_btnWrite_clicked();

    void on_btnTxtFile_clicked();
    void on_btnExcelFile_clicked();

    void on_outputUrl_editingFinished();
    void on_lineId_editingFinished();

    void on_actionAbout_triggered();

    // Helper functions
    QString writeToXlsx(QFile *txtFile, QString id, QString excelUrl, QString outputUrl);
    void readFromXlsx(QString fileUrl);


private:
    Ui::MainWindow *ui;

    bool reformatTxt(QFile *txtFile);
    bool findNextColumn(QFile *txtFile, QString id);
    QString numToAlph(int num);
};

#endif // MAINWINDOW_H
