#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "QFileDialog"

#include <QtWidgets>
#include "xlsxdocument.h"
#include "xlsxworksheet.h"
#include "xlsxcellrange.h"
#include "xlsxsheetmodel.h"

using namespace QXlsx;

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    this->ui->actionSave->setEnabled(false);
    this->ui->actionCopy->setEnabled(false);
    this->ui->actionClose_File->setEnabled(false);

    QTimer *timer = new QTimer(this);
    connect(timer, SIGNAL(timeout()), this, SLOT(update()));
    timer->start(1000);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_actionOpen_triggered()
{
    QString fileName = QFileDialog::getOpenFileName(this, tr("Open Excel File"), (m_lSheets.size() < 1) ? QString("c:/") :  QFileInfo(m_lSheets[0]->m_sXlsFilePathName).absoluteDir().absolutePath(), tr("Excel Files (*.xls *.xlsx)"));
    if (fileName.isEmpty())
    {
        return;
    }

    {
        this->ui->tabWidget->clear();
        this->setWindowTitle(fileName);
        foreach(sheet* sheet, m_lSheets)
        {
            delete sheet;
        }
        m_lSheets.clear();
    }

    QXlsx::Document* pxlsx = new QXlsx::Document(fileName);

    foreach (QString sheetName, pxlsx->sheetNames())
    {
        Worksheet *workSheet = dynamic_cast<Worksheet *>(pxlsx->sheet(sheetName));
        if (workSheet)
        {
            QTableView *view = new QTableView(this->ui->tabWidget);
            SheetModel* sheetModel = new SheetModel(workSheet, view);
            view->setModel(sheetModel);

            foreach (CellRange range, workSheet->mergedCells())
            {
                view->setSpan(range.firstRow()-1, range.firstColumn()-1, range.rowCount(), range.columnCount());
            }
            int iAddedTab = this->ui->tabWidget->addTab(view, sheetName);
            this->ui->tabWidget->setTabIcon(iAddedTab, QIcon(":/res/excel.png"));

            m_lSheets.append(new sheet(fileName, sheetName, pxlsx, sheetModel, view, workSheet, this->ui->tabWidget, iAddedTab, this));
        }
    }
    this->ui->actionSave->setEnabled(this->m_lSheets.size() > 0);
    this->ui->actionClose_File->setEnabled(this->m_lSheets.size() > 0);
}

void MainWindow::on_actionExit_triggered()
{
    this->close();
}

void MainWindow::on_actionSave_triggered()
{
    QString fileName = QFileDialog::getSaveFileName(this, tr("Save As Excel File"), QFileInfo(m_lSheets[0]->m_sXlsFilePathName).absoluteDir().absolutePath(), tr("Excel Files (*.xlsx)"));
    if (fileName.isEmpty())
    {
        return;
    }

    m_lSheets[0]->m_pXlsx->saveAs(fileName);
}

void MainWindow::keyPressEvent(QKeyEvent* event)
{
    if((this->m_lSheets.size() > 0) && (event->key() == Qt::Key_C && (event->modifiers() & Qt::ControlModifier)))
    {
        //QMessageBox::information(0, "ctrl+c","ctrl+c", 1,0,0);
        on_actionCopy_triggered();
        event->accept();
    }
}

void MainWindow::on_actionCopy_triggered()
{
    if(this->m_lSheets.size() > 0)
    {
        sheet* curSheet = this->m_lSheets[this->ui->tabWidget->currentIndex()];
        QModelIndexList cells = curSheet->m_pTableView->selectionModel()->selectedIndexes();

        QString text;
        int currentRow = 0; // To determine when to insert newlines
        foreach (const QModelIndex& cell, cells)
        {
            if (text.length() == 0)
            {
                // First item
            }
            else if (cell.row() != currentRow)
            {
                // New row
                text += '\n';
            }
            else
            {
                // Next cell
                text += '\t';
            }
            currentRow = cell.row();
            text += cell.data().toString();
        }

        if (text.length() == 0)
        {
            QMessageBox::information(0, "ctrl+c","Please select the cells to be copied!", 1,0,0);
            return;
        }

        QApplication::clipboard()->setText(text);
    }
}

void MainWindow::update()
{
    bool hasSelected = false;
    if(this->m_lSheets.size() > 0)
    {
        sheet* curSheet = this->m_lSheets[this->ui->tabWidget->currentIndex()];
        QModelIndexList cells = curSheet->m_pTableView->selectionModel()->selectedIndexes();
        hasSelected = cells.size() > 0;
    }
    this->ui->actionCopy->setEnabled(hasSelected);
}

void MainWindow::on_actionClose_File_triggered()
{
    {
        this->ui->tabWidget->clear();
        this->setWindowTitle("The Excellerator!");
        foreach(sheet* sheet, m_lSheets)
        {
            delete sheet;
        }
        m_lSheets.clear();
    }
    this->ui->actionSave->setEnabled(this->m_lSheets.size() > 0);
    this->ui->actionClose_File->setEnabled(this->m_lSheets.size() > 0);
}
