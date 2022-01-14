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

private:
    Ui::Permissible_Deviation *ui;
};
#endif // PERMISSIBLE_DEVIATION_H
