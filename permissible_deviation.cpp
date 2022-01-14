#include "permissible_deviation.h"
#include "ui_permissible_deviation.h"
#include <QGraphicsPixmapItem>

Permissible_Deviation::Permissible_Deviation(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::Permissible_Deviation)
{
    ui->setupUi(this);
    QPixmap pixmap(":/Images/Isgec-logo.jpg");
    int w = ui->label_isgec_logo->width();
    int h = ui->label_isgec_logo->height();
    ui->label_isgec_logo->setPixmap(pixmap.scaled(w,h,Qt::KeepAspectRatio));
    QImage image_292(":/Images/292.png");
    ui->lb_292->setPixmap(QPixmap::fromImage(image_292));
    QImage image_801(":/Images/801.png");
    ui->lb_801->setPixmap(QPixmap::fromImage(image_801));

}

Permissible_Deviation::~Permissible_Deviation()
{
    delete ui;
}

void Permissible_Deviation::on_le_outside_diameter_cy_shell_textChanged(const QString &arg1)
{
    float design_length = ui->le_outside_diameter_cy_shell->text().toFloat() / 2;
    ui->le_design_length->setText(QString::number(design_length));
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
    float G18 = ui->le_thickness_less_CA->text().toFloat();
    float G20 = ui->le_outside_diameter_shell->text().toFloat();
    float ratio_od_and_thk = G20/G18;
    ui->le_ratio_od_and_thk->setText(QString::number(ratio_od_and_thk));

    float G10 = ui->le_design_length->text().toFloat();
    float ratio_design_length_and_OD = G10/G20;
    ui->le_ratio_design_length_and_OD->setText(QString::number(ratio_design_length_and_OD));
}


void Permissible_Deviation::on_le_design_length_textChanged(const QString &arg1)
{
    float G10 = ui->le_design_length->text().toFloat();
    float G20 = ui->le_outside_diameter_shell->text().toFloat();
    float ratio_design_length_and_OD = G10/G20;
    ui->le_ratio_design_length_and_OD->setText(QString::number(ratio_design_length_and_OD));
}

