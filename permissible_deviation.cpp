#include "permissible_deviation.h"
#include "ui_permissible_deviation.h"

Permissible_Deviation::Permissible_Deviation(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::Permissible_Deviation)
{
    ui->setupUi(this);
}

Permissible_Deviation::~Permissible_Deviation()
{
    delete ui;
}

