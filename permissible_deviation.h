#ifndef PERMISSIBLE_DEVIATION_H
#define PERMISSIBLE_DEVIATION_H

#include <QMainWindow>

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

    void on_le_outside_diameter_cy_shell_textChanged(const QString &arg1);

    void on_le_normal_thickness_shell_editingFinished();

    void on_le_corrosion_allowance_editingFinished();

    void on_le_inside_diameter_shell_editingFinished();

    void on_le_thickness_less_CA_textChanged(const QString &arg1);

    void on_le_outside_diameter_shell_textChanged(const QString &arg1);

    void on_le_design_length_textChanged(const QString &arg1);

private:
    Ui::Permissible_Deviation *ui;
};
#endif // PERMISSIBLE_DEVIATION_H
