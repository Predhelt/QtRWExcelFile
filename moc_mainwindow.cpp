/****************************************************************************
** Meta object code from reading C++ file 'mainwindow.h'
**
** Created by: The Qt Meta Object Compiler version 67 (Qt 5.9.0)
**
** WARNING! All changes made in this file will be lost!
*****************************************************************************/

#include "../../ReadWriteExcelProj/mainwindow.h"
#include <QtCore/qbytearray.h>
#include <QtCore/qmetatype.h>
#if !defined(Q_MOC_OUTPUT_REVISION)
#error "The header file 'mainwindow.h' doesn't include <QObject>."
#elif Q_MOC_OUTPUT_REVISION != 67
#error "This file was generated using the moc from 5.9.0. It"
#error "cannot be used with the include files from this version of Qt."
#error "(The moc has changed too much.)"
#endif

QT_BEGIN_MOC_NAMESPACE
QT_WARNING_PUSH
QT_WARNING_DISABLE_DEPRECATED
struct qt_meta_stringdata_MainWindow_t {
    QByteArrayData data[18];
    char stringdata0[280];
};
#define QT_MOC_LITERAL(idx, ofs, len) \
    Q_STATIC_BYTE_ARRAY_DATA_HEADER_INITIALIZER_WITH_OFFSET(len, \
    qptrdiff(offsetof(qt_meta_stringdata_MainWindow_t, stringdata0) + ofs \
        - idx * sizeof(QByteArrayData)) \
    )
static const qt_meta_stringdata_MainWindow_t qt_meta_stringdata_MainWindow = {
    {
QT_MOC_LITERAL(0, 0, 10), // "MainWindow"
QT_MOC_LITERAL(1, 11, 26), // "on_btnReadTemplate_clicked"
QT_MOC_LITERAL(2, 38, 0), // ""
QT_MOC_LITERAL(3, 39, 24), // "on_btnReadOutput_clicked"
QT_MOC_LITERAL(4, 64, 19), // "on_btnWrite_clicked"
QT_MOC_LITERAL(5, 84, 21), // "on_btnTxtFile_clicked"
QT_MOC_LITERAL(6, 106, 23), // "on_btnExcelFile_clicked"
QT_MOC_LITERAL(7, 130, 28), // "on_outputUrl_editingFinished"
QT_MOC_LITERAL(8, 159, 25), // "on_lineId_editingFinished"
QT_MOC_LITERAL(9, 185, 24), // "on_actionAbout_triggered"
QT_MOC_LITERAL(10, 210, 11), // "writeToXlsx"
QT_MOC_LITERAL(11, 222, 6), // "QFile*"
QT_MOC_LITERAL(12, 229, 7), // "txtFile"
QT_MOC_LITERAL(13, 237, 2), // "id"
QT_MOC_LITERAL(14, 240, 8), // "excelUrl"
QT_MOC_LITERAL(15, 249, 9), // "outputUrl"
QT_MOC_LITERAL(16, 259, 12), // "readFromXlsx"
QT_MOC_LITERAL(17, 272, 7) // "fileUrl"

    },
    "MainWindow\0on_btnReadTemplate_clicked\0"
    "\0on_btnReadOutput_clicked\0on_btnWrite_clicked\0"
    "on_btnTxtFile_clicked\0on_btnExcelFile_clicked\0"
    "on_outputUrl_editingFinished\0"
    "on_lineId_editingFinished\0"
    "on_actionAbout_triggered\0writeToXlsx\0"
    "QFile*\0txtFile\0id\0excelUrl\0outputUrl\0"
    "readFromXlsx\0fileUrl"
};
#undef QT_MOC_LITERAL

static const uint qt_meta_data_MainWindow[] = {

 // content:
       7,       // revision
       0,       // classname
       0,    0, // classinfo
      10,   14, // methods
       0,    0, // properties
       0,    0, // enums/sets
       0,    0, // constructors
       0,       // flags
       0,       // signalCount

 // slots: name, argc, parameters, tag, flags
       1,    0,   64,    2, 0x08 /* Private */,
       3,    0,   65,    2, 0x08 /* Private */,
       4,    0,   66,    2, 0x08 /* Private */,
       5,    0,   67,    2, 0x08 /* Private */,
       6,    0,   68,    2, 0x08 /* Private */,
       7,    0,   69,    2, 0x08 /* Private */,
       8,    0,   70,    2, 0x08 /* Private */,
       9,    0,   71,    2, 0x08 /* Private */,
      10,    4,   72,    2, 0x08 /* Private */,
      16,    1,   81,    2, 0x08 /* Private */,

 // slots: parameters
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::Void,
    QMetaType::QString, 0x80000000 | 11, QMetaType::QString, QMetaType::QString, QMetaType::QString,   12,   13,   14,   15,
    QMetaType::Void, QMetaType::QString,   17,

       0        // eod
};

void MainWindow::qt_static_metacall(QObject *_o, QMetaObject::Call _c, int _id, void **_a)
{
    if (_c == QMetaObject::InvokeMetaMethod) {
        MainWindow *_t = static_cast<MainWindow *>(_o);
        Q_UNUSED(_t)
        switch (_id) {
        case 0: _t->on_btnReadTemplate_clicked(); break;
        case 1: _t->on_btnReadOutput_clicked(); break;
        case 2: _t->on_btnWrite_clicked(); break;
        case 3: _t->on_btnTxtFile_clicked(); break;
        case 4: _t->on_btnExcelFile_clicked(); break;
        case 5: _t->on_outputUrl_editingFinished(); break;
        case 6: _t->on_lineId_editingFinished(); break;
        case 7: _t->on_actionAbout_triggered(); break;
        case 8: { QString _r = _t->writeToXlsx((*reinterpret_cast< QFile*(*)>(_a[1])),(*reinterpret_cast< QString(*)>(_a[2])),(*reinterpret_cast< QString(*)>(_a[3])),(*reinterpret_cast< QString(*)>(_a[4])));
            if (_a[0]) *reinterpret_cast< QString*>(_a[0]) = std::move(_r); }  break;
        case 9: _t->readFromXlsx((*reinterpret_cast< QString(*)>(_a[1]))); break;
        default: ;
        }
    } else if (_c == QMetaObject::RegisterMethodArgumentMetaType) {
        switch (_id) {
        default: *reinterpret_cast<int*>(_a[0]) = -1; break;
        case 8:
            switch (*reinterpret_cast<int*>(_a[1])) {
            default: *reinterpret_cast<int*>(_a[0]) = -1; break;
            case 0:
                *reinterpret_cast<int*>(_a[0]) = qRegisterMetaType< QFile* >(); break;
            }
            break;
        }
    }
}

const QMetaObject MainWindow::staticMetaObject = {
    { &QMainWindow::staticMetaObject, qt_meta_stringdata_MainWindow.data,
      qt_meta_data_MainWindow,  qt_static_metacall, nullptr, nullptr}
};


const QMetaObject *MainWindow::metaObject() const
{
    return QObject::d_ptr->metaObject ? QObject::d_ptr->dynamicMetaObject() : &staticMetaObject;
}

void *MainWindow::qt_metacast(const char *_clname)
{
    if (!_clname) return nullptr;
    if (!strcmp(_clname, qt_meta_stringdata_MainWindow.stringdata0))
        return static_cast<void*>(const_cast< MainWindow*>(this));
    return QMainWindow::qt_metacast(_clname);
}

int MainWindow::qt_metacall(QMetaObject::Call _c, int _id, void **_a)
{
    _id = QMainWindow::qt_metacall(_c, _id, _a);
    if (_id < 0)
        return _id;
    if (_c == QMetaObject::InvokeMetaMethod) {
        if (_id < 10)
            qt_static_metacall(this, _c, _id, _a);
        _id -= 10;
    } else if (_c == QMetaObject::RegisterMethodArgumentMetaType) {
        if (_id < 10)
            qt_static_metacall(this, _c, _id, _a);
        _id -= 10;
    }
    return _id;
}
QT_WARNING_POP
QT_END_MOC_NAMESPACE
