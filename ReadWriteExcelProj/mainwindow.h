/*This project asks for the user to input the location of an existing excel file and the
desired output location and either opens the excel file to be read by the user or writes
the contents of the table to the excel file starting at A1*/
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

    QString writeToXlsx(QFile *txtFile, QString id, QString excelUrl, QString outputUrl);
    void readFromXlsx(QString fileUrl);
    bool findNextColumn(QFile *txtFile, QString id);

    void on_btnTxtFile_clicked();

    void on_btnExcelFile_clicked();

    void on_outputUrl_editingFinished();

    void on_actionAbout_triggered();

    void on_lineId_editingFinished();

private:
    Ui::MainWindow *ui;

    bool reformatTxt(QFile *txtFile);

    QString numToAlph(int num);
};

#endif // MAINWINDOW_H
