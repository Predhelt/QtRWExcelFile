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

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

public slots:
    //void exportCSV(QString fileUrl);

private slots:
    void on_btnReadTemplate_clicked();
    void on_btnReadOutput_clicked();

    void on_btnWrite_clicked();

    QString writeToXlsx(QFile *txtFile, QString id, QString excelUrl, QString outputUrl);
    bool readFromXlsx(QString fileUrl);
    bool findNextColumn(QFile *txtFile, QString id);

    void on_btnTxtFile_clicked();

    void on_btnExcelFile_clicked();

    void on_outputUrl_editingFinished();

    void on_actionAbout_triggered();

private:
    Ui::MainWindow *ui;

    void reformatTxt(QFile *txtFile);

    QString numToAlph(int num);
};

#endif // MAINWINDOW_H
