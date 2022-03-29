#include "myexcel.h"

MyExcel::MyExcel()
{

}


QVariant MyExcel::readAll()
{
    QVariant var;
    if (this->sheet != NULL && ! this->sheet->isNull())
    {
        QAxObject *usedRange = this->sheet->querySubObject("UsedRange");
        if(NULL == usedRange || usedRange->isNull())
        {
            return var;
        }
        var = usedRange->dynamicCall("Value");
        delete usedRange;
    }
    return var;
}


void MyExcel::castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res)
{
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
    {
        return;
    }
    const int rowCount = varRows.size();
    QVariantList rowData;
    for(int i=0;i<rowCount;++i)
    {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
}


/*
 *
 *  -----------------------------------
 *  Qt 下快速读写Excel指南
 *  https://blog.51cto.com/u_15329836/3402103
*/

