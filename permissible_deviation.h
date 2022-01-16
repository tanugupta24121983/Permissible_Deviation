#ifndef PERMISSIBLE_DEVIATION_H
#define PERMISSIBLE_DEVIATION_H

#include <QMainWindow>
#include <QAxObject>

QT_BEGIN_NAMESPACE
namespace Ui { class Permissible_Deviation; }
QT_END_NAMESPACE

class Permissible_Deviation : public QMainWindow
{
    Q_OBJECT

public:
    Permissible_Deviation(QWidget *parent = nullptr);
    ~Permissible_Deviation();

private slots:

    void on_le_normal_thickness_shell_editingFinished();

    void on_le_corrosion_allowance_editingFinished();

    void on_le_inside_diameter_shell_editingFinished();

    void on_le_thickness_less_CA_textChanged(const QString &arg1);

    void on_le_outside_diameter_shell_textChanged(const QString &arg1);

    void on_le_design_length_textChanged(const QString &arg1);

    void on_le_outside_diameter_shell_3_textChanged(const QString &arg1);

    void on_le_inside_diameter_shell_3_editingFinished();

    void on_le_normal_thickness_shell_3_editingFinished();

    void on_le_crown_radius_textChanged(const QString &arg1);

    void on_le_CA_4_editingFinished();

    void on_le_thickness_less_CA_3_textChanged(const QString &arg1);

    void on_le_design_length_3_textChanged(const QString &arg1);

    void on_le_nominal_thickness_DE_editingFinished();

    void on_le_max_arc_DE_textChanged(const QString &arg1);

    void on_le_Nominal_thickness_ts_textChanged(const QString &arg1);


    void on_le_max_arc_textChanged(const QString &arg1);

    void on_pb_generate_report_cylinder_clicked();

    void on_pb_calculate_max_permissible_deviation_DE_clicked();

    void on_fh_pb_generate_report_PD_DE_clicked();

    void on_pb_clear_PD_DE_clicked();

    void on_pb_calculate_max_arc_DE_clicked();

    void on_le_Max_Permissible_deviation_3_textChanged(const QString &arg1);

    void on_le_chord_length_DE_textChanged(const QString &arg1);

    void on_pb_clear_PD_CY_clicked();

    void on_pb_calculate_max_permissible_deviation_CY_clicked();

    void on_pb_calculate_max_arc_CY_clicked();

    void on_le_Max_Permissible_deviation_textChanged(const QString &arg1);

    void on_le_chord_length_cy_textChanged(const QString &arg1);

private:
    Ui::Permissible_Deviation *ui;
    QAxObject * excel;
    QAxObject * workbooks;
    QAxObject * workbook;
    QAxObject * sheets;
    QAxObject * sheet;
    float DE_80_1;
    float DE_29_2;
    float CY_80_1;
    float CY_29_2;
};
#endif // PERMISSIBLE_DEVIATION_H
