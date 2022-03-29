#ifndef MYEXCEL_H
#define MYEXCEL_H

#include <QMainWindow>
#include <QAxObject>
#include <QVariant>
class MyExcel
{
public:
    MyExcel();
    QVariant readAll();
    void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res);
    bool writeCurrentSheet(const QList<QList<QVariant> > &cells);
    void convertToColName(int data, QString &res);
    QString to26AlphabetString(int data);
private:
    QAxObject* sheet;                            //操作Excel文件对象(open-save-close-quit)
    QAxObject* workbooks;                        //总工作薄对象
    QAxObject* workbook;                         //操作当前工作薄对象
    QAxObject* worksheets;                       //文件中所有<Sheet>表页
    QAxObject* worksheet;                        //存储第n个sheet对象
    QAxObject* usedrange;                        //存储当前sheet的数据对象
};

#endif // MYEXCEL_H
