/*This project asks for the user to input the location of an existing excel file and the
desired output location and either opens the excel file to be read by the user or writes
the contents of the table to the excel file starting at A1*/
#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QFile>

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
    void on_btnRead_clicked();

    void on_btnWrite_clicked();

    bool writeToXlsx(QString fileUrl, QString outputUrl);
    void readFromXlsx(QString fileUrl);

private:
    Ui::MainWindow *ui;

    QFile *csvFile = new QFile("testExcel.csv");
    QFile *xlsxFile = new QFile("Book1.xlsx");

    QString numToAlph(int num);
    QString cellText; /* TODO: Read an input text file and retrieve the information that will go
    in each cell.  Each cell can either be read one at a time and plotted one at a time or they
    can all be read at once and plotted using specific formatting.  Figure out which one makes
    more sense or works better*/
};

#endif // MAINWINDOW_H
