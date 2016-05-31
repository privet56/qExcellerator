#include "sheet.h"

sheet::sheet(QString sXlsFilePathName, QString sSheetName, QXlsx::Document* pXlsx, SheetModel* pSheetModel, QTableView* pTableView, Worksheet* pSheet, QTabWidget* pTabWidget, int iTabIndex, QObject *parent) : QObject(parent)
{
    m_sXlsFilePathName  = sXlsFilePathName;
    m_sSheetName        = sSheetName;
    m_pXlsx             = pXlsx;
    m_pSheetModel       = pSheetModel;
    m_pTableView        = pTableView;
    m_pSheet            = pSheet;
    m_pTabWidget        = pTabWidget;
    m_iTabIndex         = iTabIndex;
}

//TODO: destructor
