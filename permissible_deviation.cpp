#include "permissible_deviation.h"
#include "ui_permissible_deviation.h"
#include <QGraphicsPixmapItem>
#include <QFile>
#include <QDebug>
#include <QAxObject>
#include <QDir>
#include <QMessageBox>


Permissible_Deviation::Permissible_Deviation(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::Permissible_Deviation),
      DE_80_1(0),
      DE_29_2(0),
      CY_80_1(0),
      CY_29_2(0)
{
    ui->setupUi(this);
    QPixmap pixmap(":/Images/Isgec-logo.jpg");
    int w = ui->label_isgec_logo->width();
    int h = ui->label_isgec_logo->height();
    ui->label_isgec_logo->setPixmap(pixmap.scaled(w,h,Qt::KeepAspectRatio));
    ui->label_isgec_logo_4->setPixmap(pixmap.scaled(w,h,Qt::KeepAspectRatio));
    QImage image_292(":/Images/292.png");
    ui->lb_292->setPixmap(QPixmap::fromImage(image_292));
    ui->lb_292_2->setPixmap(QPixmap::fromImage(image_292));
    QImage image_801(":/Images/801.png");
    ui->lb_801_2->setPixmap(QPixmap::fromImage(image_801));
    ui->lb_801->setPixmap(QPixmap::fromImage(image_801));


}

Permissible_Deviation::~Permissible_Deviation()
{
    delete ui;
}


void Permissible_Deviation::on_le_normal_thickness_shell_editingFinished()
{

    float G12 = ui->le_normal_thickness_shell->text().toFloat();
    float G16 = ui->le_corrosion_allowance->text().toFloat();
    float thickness_less_CA = G12 - G16;
    ui->le_thickness_less_CA->setText(QString::number(thickness_less_CA));
    float G14 = ui->le_inside_diameter_shell->text().toFloat();
    float outside_diameter_shell = G14 + (2* G12);
    ui->le_outside_diameter_shell->setText(QString::number(outside_diameter_shell));
    ui->le_Nominal_thickness_ts->setText(QString::number(G12));

}

void Permissible_Deviation::on_le_corrosion_allowance_editingFinished()
{

    float G12 = ui->le_normal_thickness_shell->text().toFloat();
    float G16 = ui->le_corrosion_allowance->text().toFloat();
    float thickness_less_CA = G12 - G16;
    ui->le_thickness_less_CA->setText(QString::number(thickness_less_CA));

}

void Permissible_Deviation::on_le_inside_diameter_shell_editingFinished()
{
    float G14 = ui->le_inside_diameter_shell->text().toFloat();
    float G12 = ui->le_normal_thickness_shell->text().toFloat();
    float outside_diameter_shell = G14 + (2* G12);
    ui->le_outside_diameter_shell->setText(QString::number(outside_diameter_shell));
}

void Permissible_Deviation::on_le_thickness_less_CA_textChanged(const QString &arg1)
{
    float G18 = ui->le_thickness_less_CA->text().toFloat();
    float G20 = ui->le_outside_diameter_shell->text().toFloat();
    float ratio_od_and_thk = G20/G18;
    ui->le_ratio_od_and_thk->setText(QString::number(ratio_od_and_thk));
}

void Permissible_Deviation::on_le_outside_diameter_shell_textChanged(const QString &arg1)
{
    float design_length = ui->le_outside_diameter_shell->text().toFloat() / 2;
    ui->le_design_length->setText(QString::number(design_length));

    float G18 = ui->le_thickness_less_CA->text().toFloat();
    float G20 = ui->le_outside_diameter_shell->text().toFloat();
    float ratio_od_and_thk = G20/G18;
    ui->le_ratio_od_and_thk->setText(QString::number(ratio_od_and_thk));

    float G10 = ui->le_design_length->text().toFloat();
    float ratio_design_length_and_OD = G10/G20;
    ui->le_ratio_design_length_and_OD->setText(QString::number(ratio_design_length_and_OD));


    float G31 = ui->le_Nominal_thickness_ts->text().toFloat();
    float outside_diameter_cy_shell = G20/G31;
    ui->le_outside_diameter_cy_shell->setText(QString::number(outside_diameter_cy_shell));

}

void Permissible_Deviation::on_le_Nominal_thickness_ts_textChanged(const QString &arg1)
{
    float G20 = ui->le_outside_diameter_shell->text().toFloat();
    float G31 = ui->le_Nominal_thickness_ts->text().toFloat();
    float outside_diameter_cy_shell = G20/G31;
    ui->le_outside_diameter_cy_shell->setText(QString::number(outside_diameter_cy_shell));
}

void Permissible_Deviation::on_le_design_length_textChanged(const QString &arg1)
{

    float G10 = ui->le_design_length->text().toFloat();
    float G20 = ui->le_outside_diameter_shell->text().toFloat();
    float ratio_design_length_and_OD = G10/G20;
    ui->le_ratio_design_length_and_OD->setText(QString::number(ratio_design_length_and_OD));

}

void Permissible_Deviation::on_le_max_arc_textChanged(const QString &arg1)
{
    float G36 = ui->le_max_arc->text().toDouble();
    float chord_length_cy = 2 * G36;
    ui->le_chord_length_cy->setText(QString::number(chord_length_cy));
}


void Permissible_Deviation::on_pb_generate_report_cylinder_clicked()
{
    QString fileName = "C:/ISGEC-TOOLS/OP_Template & Max Permissible Deviation For Cylinder 12.xlsx";
    QFile file_write(fileName);
    if(file_write.open(QIODevice::ReadWrite)) {
       excel     = new QAxObject("Excel.Application");
       workbooks = excel->querySubObject("Workbooks");
       workbook  = workbooks->querySubObject("Open(const QString&)",fileName);
       sheets    = workbook->querySubObject("Worksheets");
       sheet     = sheets->querySubObject("Item(int)", 1);
    }
    else {
        QMessageBox msgBox;
        msgBox.setText("Please Make Sure Excel file is closed!!");
        msgBox.exec();
        return;
    }
    file_write.close();
    auto cell = sheet->querySubObject("Cells(int,int)", 5,3);
    cell->setProperty("Value", ui->le_PD_CY_designed_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 1,6);
    cell->setProperty("Value", ui->le_PD_CY_client_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,6);
    cell->setProperty("Value", ui->le_PD_CY_eqpt_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,6);
    cell->setProperty("Value", ui->le_PD_CY_jobno_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,9);
    cell->setProperty("Value", ui->le_PD_CY_drno_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,11);
    cell->setProperty("Value", ui->le_PD_CY_rev_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 2,11);
    cell->setProperty("Value", ui->le_PD_CY_page_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 10,7);
    cell->setProperty("Value", ui->le_design_length->text());

    cell = sheet->querySubObject("Cells(int,int)", 12,7);
    cell->setProperty("Value", ui->le_normal_thickness_shell->text());

    cell = sheet->querySubObject("Cells(int,int)", 14,7);
    cell->setProperty("Value", ui->le_inside_diameter_shell->text());

    cell = sheet->querySubObject("Cells(int,int)", 16,7);
    cell->setProperty("Value", ui->le_corrosion_allowance->text());

    cell = sheet->querySubObject("Cells(int,int)", 18,7);
    cell->setProperty("Value", ui->le_thickness_less_CA->text());

    cell = sheet->querySubObject("Cells(int,int)", 20,7);
    cell->setProperty("Value", ui->le_outside_diameter_shell->text());

    cell = sheet->querySubObject("Cells(int,int)", 22,7);
    cell->setProperty("Value", ui->le_ratio_od_and_thk->text());

    cell = sheet->querySubObject("Cells(int,int)", 24,7);
    cell->setProperty("Value", ui->le_ratio_design_length_and_OD->text());

    QString factor = QString::number(CY_80_1) + "x tca ";
    cell = sheet->querySubObject("Cells(int,int)", 26,6);
    cell->setProperty("Value", factor);

    cell = sheet->querySubObject("Cells(int,int)", 27,7);
    cell->setProperty("Value", ui->le_Max_Permissible_deviation->text());

    cell = sheet->querySubObject("Cells(int,int)", 29,7);
    cell->setProperty("Value", ui->le_cy_taken->text());

    cell = sheet->querySubObject("Cells(int,int)", 31,7);
    cell->setProperty("Value", ui->le_Nominal_thickness_ts->text());

    cell = sheet->querySubObject("Cells(int,int)", 33,7);
    cell->setProperty("Value", ui->le_outside_diameter_cy_shell->text());


    factor = QString::number(CY_29_2) + "x Do ";
    cell = sheet->querySubObject("Cells(int,int)", 35,6);
    cell->setProperty("Value", factor);

    cell = sheet->querySubObject("Cells(int,int)", 36,7);
    cell->setProperty("Value", ui->le_max_arc->text());

    cell = sheet->querySubObject("Cells(int,int)", 38,7);
    cell->setProperty("Value", ui->le_chord_length_cy->text());

    cell = sheet->querySubObject("Cells(int,int)", 41,7);
    cell->setProperty("Value", ui->le_cy_taken_2->text());

    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");
    workbook->dynamicCall("Quit()");
    excel->dynamicCall("Quit()");

    file_write.close();

    QMessageBox msgBox;
    msgBox.setText("Report Has Been Succesfully Generated !!!!!");
    msgBox.exec();
    return;

}

void Permissible_Deviation::on_pb_clear_PD_CY_clicked()
{
    ui->le_design_length->clear();
    ui->le_normal_thickness_shell->clear();
    ui->le_inside_diameter_shell->clear();
    ui->le_corrosion_allowance->clear();

    ui->le_thickness_less_CA->clear();
    ui->le_outside_diameter_shell->clear();
    ui->le_ratio_od_and_thk->clear();
    ui->le_ratio_design_length_and_OD->clear();
    ui->le_Max_Permissible_deviation->clear();
    ui->le_cy_taken->clear();

    ui->le_Nominal_thickness_ts->clear();
    ui->le_outside_diameter_cy_shell->clear();
    ui->le_max_arc->clear();
    ui->le_chord_length_cy->clear();
    ui->le_cy_taken_2->clear();
}

void Permissible_Deviation::on_pb_calculate_max_permissible_deviation_CY_clicked()
{
    QString fileName = QCoreApplication::applicationDirPath() + "/CalculationSheet.xlsx";
    QFile file(fileName);
    try {
        if(file.open(QIODevice::ReadWrite)) {
            qDebug() << "File opened succesfully";
            excel     = new QAxObject("Excel.Application");
            workbooks = excel->querySubObject("Workbooks");
            workbook  = workbooks->querySubObject("Open(const QString&)", fileName);
            sheets    = workbook->querySubObject("Worksheets");
            sheet     = sheets->querySubObject("Item(int)", 3);
            file.close();
            auto cell = sheet->querySubObject("Cells(int,int)", 2,2);
            float value = ui->le_ratio_design_length_and_OD->text().toFloat();
            cell->setProperty("Value", value);
            cell = sheet->querySubObject("Cells(int,int)", 3,2);
            value = ui->le_ratio_od_and_thk->text().toFloat();
            cell->setProperty("Value", value);

            cell = sheet->querySubObject("Cells(int,int)", 5,2);
            float value_op_80_1 = cell->dynamicCall( "Value()" ).toFloat();
            CY_80_1 = value_op_80_1;
            float G18 = ui->le_thickness_less_CA->text().toFloat();
            float max_permissible_deviation = value_op_80_1 * G18;
            ui->le_Max_Permissible_deviation->setText(QString::number(max_permissible_deviation));
            workbook->dynamicCall("Save()");
            workbook->dynamicCall("Close()");
            workbook->dynamicCall("Quit()");
            excel->dynamicCall("Quit()");
        }
        else {
            QMessageBox msgBox;
            msgBox.setText("Please Make Sure " + fileName + "exists and is not open");
            msgBox.exec();
            return;
        }
        file.close();
    }
    catch (...) {
        workbook->dynamicCall("Save()");
        workbook->dynamicCall("Close()");
        workbook->dynamicCall("Quit()");
        excel->dynamicCall("Quit()");
        QMessageBox msgBox;
        file.close();
        msgBox.setText("Please Make Sure " + fileName + "is closed ");
        msgBox.exec();
        return;
    }
}

void Permissible_Deviation::on_pb_calculate_max_arc_CY_clicked()
{
    QString fileName = QCoreApplication::applicationDirPath() + "/CalculationSheet.xlsx";
    QFile file(fileName);
    try {
        if(file.open(QIODevice::ReadWrite)) {
            qDebug() << "File opened succesfully";
            excel     = new QAxObject("Excel.Application");
            workbooks = excel->querySubObject("Workbooks");
            workbook  = workbooks->querySubObject("Open(const QString&)", fileName);
            sheets    = workbook->querySubObject("Worksheets");
            sheet     = sheets->querySubObject("Item(int)", 3);
            file.close();
            auto cell = sheet->querySubObject("Cells(int,int)", 2,2);
            float value = ui->le_ratio_design_length_and_OD->text().toFloat();
            cell->setProperty("Value", value);
            cell = sheet->querySubObject("Cells(int,int)", 4,2);
            value = ui->le_outside_diameter_cy_shell->text().toFloat();
            cell->setProperty("Value", value);

            cell = sheet->querySubObject("Cells(int,int)", 6,2);
            float value_op_29_2 = cell->dynamicCall( "Value()" ).toFloat();
            CY_29_2 = value_op_29_2;
            float G20 = ui->le_outside_diameter_shell->text().toFloat();
            float max_ard_DE= value_op_29_2 * G20;
            ui->le_max_arc->setText(QString::number(max_ard_DE));
            workbook->dynamicCall("Save()");
            workbook->dynamicCall("Close()");
            workbook->dynamicCall("Quit()");
            excel->dynamicCall("Quit()");
        }
        else {
            file.close();
            QMessageBox msgBox;
            msgBox.setText("Please Make Sure " + fileName + "exists and is not open");
            msgBox.exec();
            return;
        }
    }
    catch (...) {
        workbook->dynamicCall("Save()");
        workbook->dynamicCall("Close()");
        workbook->dynamicCall("Quit()");
        excel->dynamicCall("Quit()");

        file.close();
        QMessageBox msgBox;
        msgBox.setText("Please Make Sure " + fileName + "exists and is not open");
        msgBox.exec();
        return;
    }
}

void Permissible_Deviation::on_le_Max_Permissible_deviation_textChanged(const QString &arg1)
{
    int int_value = arg1.toFloat();
    float float_value = arg1.toFloat();
    int final_value = int_value;
    if(float(int_value) != float_value) {
        final_value += 1;
    }
    ui->le_cy_taken->setText(QString::number(final_value));
}

void Permissible_Deviation::on_le_chord_length_cy_textChanged(const QString &arg1)
{
    int int_value = arg1.toFloat();
    float float_value = arg1.toFloat();
    int final_value = int_value;
    if(float(int_value) != float_value) {
        final_value += 1;
    }
    ui->le_cy_taken_2->setText(QString::number(final_value));
}

/*****************************************DE ****************************************************/

void Permissible_Deviation::on_le_outside_diameter_shell_3_textChanged(const QString &arg1)
{
    float G18 = ui->le_outside_diameter_shell_3->text().toFloat();
    float design_length_3 = ui->le_outside_diameter_shell_3->text().toDouble() * 0.5;
    ui->le_design_length_3->setText(QString::number(design_length_3));
    float G22 = ui->le_thickness_less_CA_3->text().toFloat();
    float ratio_od_and_thk_3 = G18/G22;
    ui->le_ratio_od_and_thk_3->setText(QString::number(ratio_od_and_thk_3));

    float G33 = ui->le_nominal_thickness_DE->text().toFloat();
    float outside_diameter_cy_shell_3 = G18/G33;
    ui->le_outside_diameter_cy_shell_3->setText(QString::number(outside_diameter_cy_shell_3));
}

void Permissible_Deviation::on_le_inside_diameter_shell_3_editingFinished()
{
    float D14 = ui->le_inside_diameter_shell_3->text().toFloat();
    float crown_radius = 0.9 * D14;
    ui->le_crown_radius->setText(QString::number(crown_radius));
}

void Permissible_Deviation::on_le_normal_thickness_shell_3_editingFinished()
{
    float G12 = ui->le_normal_thickness_shell_3->text().toFloat();
    float G16 = ui->le_crown_radius->text().toFloat();
    float outside_diameter_shell_3 = 2*(G16+G12);
    ui->le_outside_diameter_shell_3->setText(QString::number(outside_diameter_shell_3));
    float G20 = ui->le_CA_4->text().toFloat();
    float thickness_less_CA_3 = G12-G20;
    ui->le_thickness_less_CA_3->setText(QString::number(thickness_less_CA_3));
    ui->le_nominal_thickness_DE->setText(QString::number(G12));
}

void Permissible_Deviation::on_le_crown_radius_textChanged(const QString &arg1)
{
    float G12 = ui->le_normal_thickness_shell_3->text().toFloat();
    float G16 = ui->le_crown_radius->text().toFloat();
    float outside_diameter_shell_3 = 2*(G16+G12);
    ui->le_outside_diameter_shell_3->setText(QString::number(outside_diameter_shell_3));
}

void Permissible_Deviation::on_le_CA_4_editingFinished()
{
    float G12 = ui->le_normal_thickness_shell_3->text().toFloat();
    float G20 = ui->le_CA_4->text().toFloat();
    float thickness_less_CA_3 = G12-G20;
    ui->le_thickness_less_CA_3->setText(QString::number(thickness_less_CA_3));
}


void Permissible_Deviation::on_le_thickness_less_CA_3_textChanged(const QString &arg1)
{
    float G18 = ui->le_outside_diameter_shell_3->text().toFloat();
    float G22 = ui->le_thickness_less_CA_3->text().toFloat();
    float ratio_od_and_thk_3 = G18/G22;
    ui->le_ratio_od_and_thk_3->setText(QString::number(ratio_od_and_thk_3));
    float G10 = ui->le_design_length_3->text().toDouble();
    float ratio_design_length_and_OD_3 = G10/G18;
    ui->le_ratio_design_length_and_OD_3->setText(QString::number(ratio_design_length_and_OD_3));

}


void Permissible_Deviation::on_le_design_length_3_textChanged(const QString &arg1)
{
    float G10 = ui->le_design_length_3->text().toDouble();
    float G18 = ui->le_outside_diameter_shell_3->text().toFloat();
    float ratio_design_length_and_OD_3 = G10/G18;
    ui->le_ratio_design_length_and_OD_3->setText(QString::number(ratio_design_length_and_OD_3));
}


void Permissible_Deviation::on_le_nominal_thickness_DE_editingFinished()
{
    float G33 = ui->le_nominal_thickness_DE->text().toFloat();
    float G18 = ui->le_outside_diameter_shell_3->text().toFloat();
    float outside_diameter_cy_shell_3 = G18/G33;
    ui->le_outside_diameter_cy_shell_3->setText(QString::number(outside_diameter_cy_shell_3));
}


void Permissible_Deviation::on_le_max_arc_DE_textChanged(const QString &arg1)
{
    float arc = ui->le_max_arc_DE->text().toFloat();
    float chord_length_DE = 2 * arc;
    ui->le_chord_length_DE->setText(QString::number(chord_length_DE));
}

void Permissible_Deviation::on_pb_calculate_max_permissible_deviation_DE_clicked()
{
    QString fileName = QCoreApplication::applicationDirPath() + "/CalculationSheet.xlsx";
    QFile file(fileName);
    try {
        if(file.open(QIODevice::ReadWrite)) {
            qDebug() << "File opened succesfully";
            excel     = new QAxObject("Excel.Application");
            workbooks = excel->querySubObject("Workbooks");
            workbook  = workbooks->querySubObject("Open(const QString&)", fileName);
            sheets    = workbook->querySubObject("Worksheets");
            sheet     = sheets->querySubObject("Item(int)", 3);
            file.close();
            auto cell = sheet->querySubObject("Cells(int,int)", 2,2);
            float value = ui->le_ratio_design_length_and_OD_3->text().toFloat();
            cell->setProperty("Value", value);
            cell = sheet->querySubObject("Cells(int,int)", 3,2);
            value = ui->le_ratio_od_and_thk_3->text().toFloat();
            cell->setProperty("Value", value);

            cell = sheet->querySubObject("Cells(int,int)", 5,2);
            float value_op_80_1 = cell->dynamicCall( "Value()" ).toFloat();
            DE_80_1 = value_op_80_1;
            float G22 = ui->le_thickness_less_CA_3->text().toFloat();
            float max_permissible_deviation = value_op_80_1 * G22;
            ui->le_Max_Permissible_deviation_3->setText(QString::number(max_permissible_deviation));
            workbook->dynamicCall("Save()");
            workbook->dynamicCall("Close()");
            workbook->dynamicCall("Quit()");
            excel->dynamicCall("Quit()");
        }
        else {
            file.close();
            QMessageBox msgBox;
            msgBox.setText("Please Make Sure " + fileName + "exists and is not open");
            msgBox.exec();
            return;
        }
    }
    catch (...) {
        workbook->dynamicCall("Save()");
        workbook->dynamicCall("Close()");
        workbook->dynamicCall("Quit()");
        excel->dynamicCall("Quit()");
        QMessageBox msgBox;
        file.close();
        msgBox.setText("Please Make Sure " + fileName + "is closed ");
        msgBox.exec();
        return;
    }

}

void Permissible_Deviation::on_pb_calculate_max_arc_DE_clicked()
{
    QString fileName = QCoreApplication::applicationDirPath() + "/CalculationSheet.xlsx";
    QFile file(fileName);
    try {
        if(file.open(QIODevice::ReadWrite)) {
            qDebug() << "File opened succesfully";
            excel     = new QAxObject("Excel.Application");
            workbooks = excel->querySubObject("Workbooks");
            workbook  = workbooks->querySubObject("Open(const QString&)", fileName);
            sheets    = workbook->querySubObject("Worksheets");
            sheet     = sheets->querySubObject("Item(int)", 3);
            file.close();
            auto cell = sheet->querySubObject("Cells(int,int)", 2,2);
            float value = ui->le_ratio_design_length_and_OD_3->text().toFloat();
            cell->setProperty("Value", value);
            cell = sheet->querySubObject("Cells(int,int)", 4,2);
            value = ui->le_outside_diameter_cy_shell_3->text().toFloat();
            cell->setProperty("Value", value);

            cell = sheet->querySubObject("Cells(int,int)", 6,2);
            float value_op_29_2 = cell->dynamicCall( "Value()" ).toFloat();
            DE_29_2 = value_op_29_2;
            float G18 = ui->le_outside_diameter_shell_3->text().toFloat();
            float max_ard_DE= value_op_29_2 * G18;
            ui->le_max_arc_DE->setText(QString::number(max_ard_DE));
            workbook->dynamicCall("Save()");
            workbook->dynamicCall("Close()");
            workbook->dynamicCall("Quit()");
            excel->dynamicCall("Quit()");
        }
        else {
            file.close();
            QMessageBox msgBox;
            msgBox.setText("Please Make Sure " + fileName + "exists and is not open");
            msgBox.exec();
            return;
        }
    }
    catch (...) {
        workbook->dynamicCall("Save()");
        workbook->dynamicCall("Close()");
        workbook->dynamicCall("Quit()");
        excel->dynamicCall("Quit()");

        file.close();
        QMessageBox msgBox;
        msgBox.setText("Please Make Sure " + fileName + "is closed ");
        msgBox.exec();
        return;
    }
}


void Permissible_Deviation::on_fh_pb_generate_report_PD_DE_clicked()
{
    QString fileName = ":/Images/Book1.xlsx";
    QFile file_write(fileName);
    if(file_write.open(QIODevice::ReadWrite)) {
       excel     = new QAxObject("Excel.Application");
       workbooks = excel->querySubObject("Workbooks");
       workbook  = workbooks->querySubObject("Open(const QString&)",fileName);
       sheets    = workbook->querySubObject("Worksheets");
       sheet     = sheets->querySubObject("Item(int)", 1);
    }
    else {
        QMessageBox msgBox;
        msgBox.setText("Please Make Sure Excel file is closed!!");
        msgBox.exec();
        return;
    }
    file_write.close();
    auto cell = sheet->querySubObject("Cells(int,int)", 5,3);
    cell->setProperty("Value", ui->le_PD_DE_designed_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 1,6);
    cell->setProperty("Value", ui->le_PD_DE_client_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 3,6);
    cell->setProperty("Value", ui->le_PD_DE_eqpt_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,6);
    cell->setProperty("Value", ui->le_PD_DE_jobno_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,9);
    cell->setProperty("Value", ui->le_PD_DE_drno_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 5,11);
    cell->setProperty("Value", ui->le_PD_DE_rev_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 2,11);
    cell->setProperty("Value", ui->le_PD_DE_page_by->text());

    cell = sheet->querySubObject("Cells(int,int)", 10,7);
    cell->setProperty("Value", ui->le_design_length_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 12,7);
    cell->setProperty("Value", ui->le_normal_thickness_shell_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 14,7);
    cell->setProperty("Value", ui->le_inside_diameter_shell_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 16,7);
    cell->setProperty("Value", ui->le_crown_radius->text());

    cell = sheet->querySubObject("Cells(int,int)", 18,7);
    cell->setProperty("Value", ui->le_outside_diameter_shell_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 20,7);
    cell->setProperty("Value", ui->le_CA_4->text());

    cell = sheet->querySubObject("Cells(int,int)", 22,7);
    cell->setProperty("Value", ui->le_thickness_less_CA_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 24,7);
    cell->setProperty("Value", ui->le_ratio_od_and_thk_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 26,7);
    cell->setProperty("Value", ui->le_ratio_design_length_and_OD_3->text());

    QString factor = QString::number(DE_80_1) + "x tca ";
    cell = sheet->querySubObject("Cells(int,int)", 28,6);
    cell->setProperty("Value", factor);

    cell = sheet->querySubObject("Cells(int,int)", 29,7);
    cell->setProperty("Value", ui->le_Max_Permissible_deviation_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 31,7);
    cell->setProperty("Value", ui->le_taken_DE_1->text());

    cell = sheet->querySubObject("Cells(int,int)", 33,7);
    cell->setProperty("Value", ui->le_nominal_thickness_DE->text());

    cell = sheet->querySubObject("Cells(int,int)", 35,7);
    cell->setProperty("Value", ui->le_outside_diameter_cy_shell_3->text());

    cell = sheet->querySubObject("Cells(int,int)", 35,7);
    cell->setProperty("Value", ui->le_outside_diameter_cy_shell_3->text());

    factor = QString::number(DE_29_2) + "x Do ";
    cell = sheet->querySubObject("Cells(int,int)", 37,6);
    cell->setProperty("Value", factor);

    cell = sheet->querySubObject("Cells(int,int)", 38,7);
    cell->setProperty("Value", ui->le_max_arc_DE->text());

    cell = sheet->querySubObject("Cells(int,int)", 40,7);
    cell->setProperty("Value", ui->le_chord_length_DE->text());

    cell = sheet->querySubObject("Cells(int,int)", 43,7);
    cell->setProperty("Value", ui->le_taken_DE_2->text());

    workbook->dynamicCall("Save()");
    workbook->dynamicCall("Close()");
    workbook->dynamicCall("Quit()");
    excel->dynamicCall("Quit()");

    file_write.close();

    QMessageBox msgBox;
    msgBox.setText("Report Has Been Succesfully Generated !!!!!");
    msgBox.exec();
    return;

}


void Permissible_Deviation::on_pb_clear_PD_DE_clicked()
{
    ui->le_design_length_3->clear();
    ui->le_normal_thickness_shell_3->clear();
    ui->le_inside_diameter_shell_3->clear();
    ui->le_crown_radius->clear();

    ui->le_outside_diameter_shell_3->clear();
    ui->le_CA_4->clear();
    ui->le_thickness_less_CA_3->clear();
    ui->le_ratio_od_and_thk_3->clear();
    ui->le_ratio_design_length_and_OD_3->clear();
    ui->le_Max_Permissible_deviation_3->clear();
    ui->le_taken_DE_1->clear();

    ui->le_nominal_thickness_DE->clear();
    ui->le_outside_diameter_cy_shell_3->clear();
    ui->le_max_arc_DE->clear();
    ui->le_chord_length_DE->clear();
    ui->le_taken_DE_2->clear();
}



void Permissible_Deviation::on_le_Max_Permissible_deviation_3_textChanged(const QString &arg1)
{
    int int_value = arg1.toFloat();
    float float_value = arg1.toFloat();
    int final_value = int_value;
    if(float(int_value) != float_value) {
        final_value += 1;
    }
    ui->le_taken_DE_1->setText(QString::number(final_value));
}


void Permissible_Deviation::on_le_chord_length_DE_textChanged(const QString &arg1)
{
    int int_value = arg1.toFloat();
    float float_value = arg1.toFloat();
    int final_value = int_value;
    if(float(int_value) != float_value) {
        final_value += 1;
    }
    ui->le_taken_DE_2->setText(QString::number(final_value));
}
