#include "widget.h"
#include "ui_widget.h"
#include <QAxObject>
#include <QFileDialog>
#include <QDebug>
#include <QStandardPaths>
#include <QDateTime>
#include <cmath>

Widget::Widget(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::Widget)
{
    ui->setupUi(this);

    //设置表格多少列，行
    ui->tableWidget->setRowCount(8);
    ui->tableWidget->setColumnCount(26);

    ui->tableWidget_2->setRowCount(5);
    ui->tableWidget_2->setColumnCount(5);

    //给表头设置边框
    ui->tableWidget->horizontalHeader()->setStyleSheet("QHeaderView::section{padding-left: 4px;border: 1px solid #383838;}");
    ui->tableWidget_2->horizontalHeader()->setStyleSheet("QHeaderView::section{padding-left: 4px;border: 1px solid #383838;}");

    //给表2设置表头
    QStringList list;
    list << "姓名" << "缺勤时间" << "加班时间"<<"迟到次数"<<"早退次数";
        ui->tableWidget_2->setHorizontalHeaderLabels(list);

    connect(ui->inert_table,&QPushButton::clicked,this,&Widget::insert_table);
    connect(ui->start_suan,&QPushButton::clicked, this, &Widget::calculate_result);

    ui->tableWidget->setSpan(0, 0, 2, 1);
    ui->tableWidget->setSpan(0, 1, 2, 1);
    ui->tableWidget->setSpan(0, 2, 2, 1);
    ui->tableWidget->setSpan(0, 3, 1, 4);
     ui->tableWidget->setSpan(0, 7, 1, 4);
      ui->tableWidget->setSpan(0, 11, 1, 5);
       ui->tableWidget->setSpan(0, 16, 1, 5);




}

Widget::~Widget()
{
    delete ui;
}

void Widget::insert_table()
{
    QAxObject *excel = new QAxObject(this);
        excel->setControl("Excel.Application");
        excel->setProperty("Visible", false);    //显示窗体看效果,选择ture将会看到excel表格被打开
        excel->setProperty("DisplayAlerts", true);
        QAxObject *workbooks = excel->querySubObject("WorkBooks");   //获取工作簿(excel文件)集合

        QString filepath = QFileDialog::getOpenFileName(this,"打开",
                                                   QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation),
                                                   "Excel 文件(*.xls *.xlsx *.xlsm)");

        //打开刚才选定的excel
        workbooks->dynamicCall("Open(const QString&)", filepath);
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");
        QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)",1);
        QAxObject *usedRange = worksheet->querySubObject("UsedRange");   //获取表格中的数据范围

        QVariant var = usedRange->dynamicCall("Value");  //将所有的数据读取刀QVariant容器中保存
        QList<QList<QVariant>> excel_list;  //用于将QVariant转换为Qlist的二维数组
        QVariantList varRows=var.toList();
        if(varRows.isEmpty())
        {
             return;
        }

        const int row_count = varRows.size();
        QVariantList rowData;
        for(int i=0;i<row_count;++i)
        {
            rowData = varRows[i].toList();
            excel_list.push_back(rowData);
        }

        qDebug()<<"row::"<<row_count;

        //打印excel数据
        for(int i = 2; i<row_count; i++)
        {
            QList<QVariant> curList = excel_list.at(i);
            int curRowCount = curList.size();
            for(int j = 0; j < curRowCount; j++)
            {
                //qDebug() << curList.at(j).toString();

                QString curstr = curList.at(j).toString();


                if(curstr.contains("T00:00:00.000", Qt::CaseSensitive))
                {
                    QString currenttime = curstr.mid(0,10);

                    //qDebug()<<"time:"<<currenttime;

                    ui->tableWidget->setItem(i-2,j,new QTableWidgetItem(currenttime));
                }
                else if(curstr.contains("0.", Qt::CaseSensitive))
                {

                    double tmpnum = curstr.toDouble();
                    allsecond =  tmpnum * 24 * 60 * 60 + 10;
                    //qDebug()<<"allsecond:"<<allsecond;

                    QTime a(0,0);
                    a = a.addSecs(int(allsecond));

                    QString strtime = a.toString("hh:mm");
                    //qDebug()<<"str:"<<strtime;

                    ui->tableWidget->setItem(i-2,j,new QTableWidgetItem(strtime));
                }
                else
                {
                    ui->tableWidget->setItem(i-2,j,new QTableWidgetItem(curList.at(j).toString()));

                }
            }
        }

        ui->tableWidget->item(0,3)->setTextAlignment(Qt::AlignHCenter | Qt::AlignVCenter);
        ui->tableWidget->item(0,7)->setTextAlignment(Qt::AlignHCenter | Qt::AlignVCenter);
        ui->tableWidget->item(0,11)->setTextAlignment(Qt::AlignHCenter | Qt::AlignVCenter);
        ui->tableWidget->item(0,16)->setTextAlignment(Qt::AlignHCenter | Qt::AlignVCenter);


        workbook->dynamicCall( "Close(Boolean)", false );
        excel->dynamicCall( "Quit(void)" );
        delete excel;
}

void Widget::calculate_result()
{
    int i = 0;
    int j = 0;
    int m = 0;
    int n = 1;

    for(i = 0; i < 5; i++)
    {
        ui->tableWidget_2->setItem(i,0, new QTableWidgetItem(ui->tableWidget->item(i+2,2)->text()));
    }

    for(i = 2; i < 7; ++i)
    {
        for(j = 4; j < 5; ++j)
        {
            qDebug()<<"time:"<<ui->tableWidget->item(i,j)->text();

            QDateTime t1time = QDateTime::fromString(ui->tableWidget->item(0,3)->text(),"yyyy-MM-dd");
            QString t1week = t1time.toString("ddd");
            if(t1week == "周六" || t1week == "周日")
            {
                if(ui->tableWidget->item(i,j)->text() == NULL && ui->tableWidget->item(i,j+1)->text() == NULL)
                {
                    queqintime+=28800;
                }
                else
                {
                    chidaotime = 0;
                    zaotuitime = 0;

                    QString time_str = ui->tableWidget->item(i,j)->text();
                    QString time_str2 = ui->tableWidget->item(i,j+1)->text();

                    QTime time=QTime::fromString(time_str);
                    int allsecond = time.hour()*60*60+time.minute()*60;

                    QTime time2=QTime::fromString(time_str2);
                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;

                    jiabantime = allsecond2 - allsecond;
                }
            }
            else
            {
                if(ui->tableWidget->item(i,j)->text() == NULL)
                {
                    queqintime+=28800;
                }
                else
                {
                    //迟到
                    QString time_str = ui->tableWidget->item(i,j)->text(); //时分秒
                    QTime time=QTime::fromString(time_str);

                    int allsecond = time.hour()*60*60+time.minute()*60;

                 //   qDebug()<<"1:"<<allsecond;

                    QString time_str2 = "09:10:00";
                    QTime time2=QTime::fromString(time_str2);

                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;
                  //  qDebug()<<"2:"<<allsecond2;

                    //早退
                    //qDebug()<<" yi "<<ui->tableWidget->item(i,j+1)->text();
                    QString time_str3 = ui->tableWidget->item(i,j+1)->text(); //时分秒
                    QTime time3=QTime::fromString(time_str3);

                    int allsecond3 = time3.hour()*60*60+time3.minute()*60;
                   // qDebug()<<"3:"<<allsecond3;


                    QString time_str4 = "18:00:00";
                    QTime time4=QTime::fromString(time_str4);

                    int allsecond4 = time4.hour()*60*60+time4.minute()*60;
                   // qDebug()<<"4:"<<allsecond4;

                    //加班
                    QString time_str5 = "19:00";
                    QTime time5=QTime::fromString(time_str5);

                    int allsecond5 = time5.hour()*60*60+time5.minute()*60;

                    if(allsecond > allsecond2)
                    {
                        chidaotime++;
                    }
                    if(allsecond3 < allsecond4)
                    {
                        zaotuitime++;
                    }
                    if(allsecond3 > allsecond5)
                    {
                        jiabantime = allsecond3 - allsecond5;
                    }
                }

            }

        }

        QTime a(0,0);
        a = a.addSecs(int(queqintime));

        QString strtime = a.toString("hh:mm");

        ui->tableWidget_2->setItem(m, n, new QTableWidgetItem(strtime));
        ui->tableWidget_2->setItem(m, n+1, new QTableWidgetItem(QString::number((jiabantime))));
        ui->tableWidget_2->setItem(m, n+2, new QTableWidgetItem(QString::number((chidaotime))));
        ui->tableWidget_2->setItem(m, n+3, new QTableWidgetItem(QString::number((zaotuitime))));

        m++;

        queqintime = 0;

        jiabantime = 0;

        chidaotime = 0;

        zaotuitime = 0;

    }


    m = 0;
    n = 1;
    queqintime = 0;

    jiabantime = 0;

    chidaotime = 0;

    zaotuitime = 0;

    for(i = 2; i < 7; ++i)
    {
        for(j = 7; j < 8; ++j)
        {
            qDebug()<<"time:"<<ui->tableWidget->item(i,j)->text();

            QDateTime t1time = QDateTime::fromString(ui->tableWidget->item(0,7)->text(),"yyyy-MM-dd");
            QString t1week = t1time.toString("ddd");
            if(t1week == "周六" || t1week == "周日")
            {
                if(ui->tableWidget->item(i,j)->text() == NULL)
                {
                    queqintime+=28800;
                    jiabantime =0;
                }
                else
                {

                    QString time_str = ui->tableWidget->item(i,j)->text();
                    //qDebug()<<" time_str "<<time_str;

                    QString time_str2 = ui->tableWidget->item(i,j+1)->text();
                    //qDebug()<<" time_str2 "<<time_str2;

                    QTime time=QTime::fromString(time_str);
                    int allsecond = time.hour()*60*60+time.minute()*60;
                    //qDebug()<<" allsecond "<<allsecond;

                    QTime time2=QTime::fromString(time_str2);
                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;
                    //qDebug()<<" allsecond2 "<<allsecond2;

                    jiabantime = allsecond2 - allsecond;
                }
            }
            else
            {
                if(ui->tableWidget->item(i,j)->text() == NULL)
                {
                    queqintime+=28800;
                }
                else
                {
                    //迟到
                    QString time_str = ui->tableWidget->item(i,j)->text(); //时分秒
                    QTime time=QTime::fromString(time_str);

                    int allsecond = time.hour()*60*60+time.minute()*60;

                   // qDebug()<<"1:"<<allsecond;

                    QString time_str2 = "09:10:00";
                    QTime time2=QTime::fromString(time_str2);

                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;
                   // qDebug()<<"2:"<<allsecond2;

                    //早退
                    //qDebug()<<" yi "<<ui->tableWidget->item(i,j+1)->text();
                    QString time_str3 = ui->tableWidget->item(i,j+1)->text(); //时分秒
                    QTime time3=QTime::fromString(time_str3);

                    int allsecond3 = time3.hour()*60*60+time3.minute()*60;
                   // qDebug()<<"3:"<<allsecond3;


                    QString time_str4 = "18:00:00";
                    QTime time4=QTime::fromString(time_str4);

                    int allsecond4 = time4.hour()*60*60+time4.minute()*60;
                    //qDebug()<<"4:"<<allsecond4;

                    //加班
                    QString time_str5 = "19:00";
                    QTime time5=QTime::fromString(time_str5);

                    int allsecond5 = time5.hour()*60*60+time5.minute()*60;

                    if(allsecond > allsecond2)
                    {
                        chidaotime++;
                    }
                    if(allsecond3 < allsecond4)
                    {
                        zaotuitime++;
                    }
                    if(allsecond3 > allsecond5)
                    {
                        jiabantime = allsecond3 - allsecond5;
                    }
                }

            }

        }

        //把表格里的数据转成数字，进行相加
        //转换缺勤
        QString time_str = ui->tableWidget_2->item(m,n)->text(); //时分秒
        QTime time=QTime::fromString(time_str);

        int allsecond = time.hour()*60*60+time.minute()*60;
        allsecond += queqintime;

        QTime a(0,0);
        a = a.addSecs(int(allsecond));

        QString strtime = a.toString("hh:mm");
        //转换加班
        int time_str2 = ui->tableWidget_2->item(m,n+1)->text().toInt(); //时分
        jiabantime +=time_str2;

        QTime a2(0,0);
        a2 = a2.addSecs(int(jiabantime));

        QString strtime2 = a2.toString("hh:mm");

        //转换迟到
        int time_str3 = ui->tableWidget_2->item(m,n+2)->text().toInt();

        time_str3 += chidaotime;

        //转换早退
        int time_str4 = ui->tableWidget_2->item(m,n+3)->text().toInt();
        time_str4 += zaotuitime;

        ui->tableWidget_2->setItem(m, n, new QTableWidgetItem(strtime));
        ui->tableWidget_2->setItem(m, n+1, new QTableWidgetItem(strtime2));
        ui->tableWidget_2->setItem(m, n+2, new QTableWidgetItem(QString::number((time_str3))));
        ui->tableWidget_2->setItem(m, n+3, new QTableWidgetItem(QString::number((time_str4))));

        m++;

        queqintime = 0;

        jiabantime = 0;

        chidaotime = 0;

        zaotuitime = 0;

    }

    m = 0;
    n = 1;
    queqintime = 0;

    jiabantime = 0;

    chidaotime = 0;

    zaotuitime = 0;

    for(i = 2; i < 7; ++i)
    {
        for(j = 12; j < 13; ++j)
        {
            qDebug()<<"time:"<<ui->tableWidget->item(i,j)->text();

            QDateTime t1time = QDateTime::fromString(ui->tableWidget->item(0,11)->text(),"yyyy-MM-dd");
            QString t1week = t1time.toString("ddd");
            if(t1week == "周六" || t1week == "周日")
            {
                if(ui->tableWidget->item(i,j)->text() == NULL)
                {
                    queqintime+=28800;
                    jiabantime =0;
                }
                else
                {

                    QString time_str = ui->tableWidget->item(i,j)->text();
                  //  qDebug()<<" time_str "<<time_str;

                    QString time_str2 = ui->tableWidget->item(i,j+1)->text();
                   // qDebug()<<" time_str2 "<<time_str2;

                    QTime time=QTime::fromString(time_str);
                    int allsecond = time.hour()*60*60+time.minute()*60;
                    //qDebug()<<" allsecond "<<allsecond;

                    QTime time2=QTime::fromString(time_str2);
                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;
                    //qDebug()<<" allsecond2 "<<allsecond2;

                    jiabantime = allsecond2 - allsecond;
                }
            }
            else
            {
                if(ui->tableWidget->item(i,j)->text() == NULL)
                {
                    queqintime+=28800;
                }
                else
                {
                    //迟到
                    QString time_str = ui->tableWidget->item(i,j)->text(); //时分秒
                    QTime time=QTime::fromString(time_str);

                    int allsecond = time.hour()*60*60+time.minute()*60;

                   // qDebug()<<"1:"<<allsecond;

                    QString time_str2 = "09:10:00";
                    QTime time2=QTime::fromString(time_str2);

                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;
                   // qDebug()<<"2:"<<allsecond2;

                    //早退
                    //qDebug()<<" yi "<<ui->tableWidget->item(i,j+1)->text();
                    QString time_str3 = ui->tableWidget->item(i,j+1)->text(); //时分秒
                    QTime time3=QTime::fromString(time_str3);

                    int allsecond3 = time3.hour()*60*60+time3.minute()*60;
                   // qDebug()<<"3:"<<allsecond3;


                    QString time_str4 = "18:00:00";
                    QTime time4=QTime::fromString(time_str4);

                    int allsecond4 = time4.hour()*60*60+time4.minute()*60;
                    //qDebug()<<"4:"<<allsecond4;

                    //加班
                    QString time_str5 = "19:00";
                    QTime time5=QTime::fromString(time_str5);

                    int allsecond5 = time5.hour()*60*60+time5.minute()*60;

                    if(allsecond > allsecond2)
                    {
                        chidaotime++;
                    }
                    if(allsecond3 < allsecond4)
                    {
                        zaotuitime++;
                    }
                    if(allsecond3 > allsecond5)
                    {
                        jiabantime = allsecond3 - allsecond5;
                    }
                }

            }

        }

        //把表格里的数据转成数字，进行相加
        //转换缺勤
        QString time_str = ui->tableWidget_2->item(m,n)->text(); //时分秒

        QTime time=QTime::fromString(time_str);

        int allsecond = time.hour()*60*60+time.minute()*60;

        allsecond += queqintime;

        QString strtime;

        if(allsecond >= 86400)
        {
            int num = allsecond/60/60;

            strtime = QString::number(num) +":"+"00";
        }
        else
        {
            QTime a(0,0);
            a = a.addSecs(int(allsecond));

            strtime = a.toString("hh:mm");
        }


        //转换加班

        QString strtime2;
        QString jia_str = ui->tableWidget_2->item(m,n+1)->text(); //时分秒
        qDebug()<<"jiajia"<<jia_str;

        QTime jia_time=QTime::fromString(jia_str);

        int jia_second = jia_time.hour()*60*60+jia_time.minute()*60;

        qDebug()<<"secon"<<jia_second;

        if(jia_second > 0)
        {
            jiabantime +=jia_second;
            QTime a(0,0);
            a = a.addSecs(int(jiabantime));

            strtime2 = a.toString("hh:mm");

        }
        else
        {

            jiabantime +=0;

            QTime a2(0,0);
            a2 = a2.addSecs(int(jiabantime));

            strtime2 = a2.toString("hh:mm");
        }


        //转换迟到
        int time_str3 = ui->tableWidget_2->item(m,n+2)->text().toInt();

        time_str3 += chidaotime;

        //转换早退
        int time_str4 = ui->tableWidget_2->item(m,n+3)->text().toInt();
        time_str4 += zaotuitime;

        ui->tableWidget_2->setItem(m, n, new QTableWidgetItem(strtime));
        ui->tableWidget_2->setItem(m, n+1, new QTableWidgetItem(strtime2));
        ui->tableWidget_2->setItem(m, n+2, new QTableWidgetItem(QString::number((time_str3))));
        ui->tableWidget_2->setItem(m, n+3, new QTableWidgetItem(QString::number((time_str4))));

        m++;

        queqintime = 0;

        jiabantime = 0;

        chidaotime = 0;

        zaotuitime = 0;

    }

    m = 0;
    n = 1;
    queqintime = 0;

    jiabantime = 0;

    chidaotime = 0;

    zaotuitime = 0;

    QString jia_str1 = ui->tableWidget_2->item(m,n+1)->text(); //时分秒
    qDebug()<<"diyigejiaban"<<jia_str1;

    for(i = 2; i < 7; ++i)
    {
        for(j = 17; j < 18; ++j)
        {
            qDebug()<<"time:"<<ui->tableWidget->item(i,j)->text();

            QDateTime t1time = QDateTime::fromString(ui->tableWidget->item(0,16)->text(),"yyyy-MM-dd");
            QString t1week = t1time.toString("ddd");
            if(t1week == "周六" || t1week == "周日")
            {
                if(ui->tableWidget->item(i,j)->text() == NULL)
                {
                    queqintime+=28800;
                    jiabantime =0;
                }
                else
                {

                    QString time_str = ui->tableWidget->item(i,j)->text();
                  //  qDebug()<<" time_str "<<time_str;

                    QString time_str2 = ui->tableWidget->item(i,j+1)->text();
                   // qDebug()<<" time_str2 "<<time_str2;

                    QTime time=QTime::fromString(time_str);
                    int allsecond = time.hour()*60*60+time.minute()*60;
                    //qDebug()<<" allsecond "<<allsecond;

                    QTime time2=QTime::fromString(time_str2);
                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;
                    //qDebug()<<" allsecond2 "<<allsecond2;

                    jiabantime = allsecond2 - allsecond;
                }
            }
            else
            {
                if(ui->tableWidget->item(i,j)->text() == NULL)
                {
                    queqintime+=28800;
                }
                else
                {
                    //迟到
                    QString time_str = ui->tableWidget->item(i,j)->text(); //时分秒
                    QTime time=QTime::fromString(time_str);

                    int allsecond = time.hour()*60*60+time.minute()*60;

                    qDebug()<<"1:"<<allsecond;

                    QString time_str2 = "09:10:00";
                    QTime time2=QTime::fromString(time_str2);

                    int allsecond2 = time2.hour()*60*60+time2.minute()*60;
                    qDebug()<<"2:"<<allsecond2;

                    //早退
                    //qDebug()<<" yi "<<ui->tableWidget->item(i,j+1)->text();
                    QString time_str3 = ui->tableWidget->item(i,j+1)->text(); //时分秒

                    qDebug()<<"lie shu:"<<time_str3;

                    QTime time3=QTime::fromString(time_str3);

                    int allsecond3 = time3.hour()*60*60+time3.minute()*60;
                   // qDebug()<<"3:"<<allsecond3;


                    QString time_str4 = "18:00:00";
                    QTime time4=QTime::fromString(time_str4);

                    int allsecond4 = time4.hour()*60*60+time4.minute()*60;
                    //qDebug()<<"4:"<<allsecond4;

                    //加班
                    QString time_str5 = "19:00";
                    QTime time5=QTime::fromString(time_str5);

                    int allsecond5 = time5.hour()*60*60+time5.minute()*60;

                    if(allsecond > allsecond2)
                    {
                        chidaotime++;
                        qDebug()<<"chidao"<<chidaotime;

                    }
                    if(allsecond3 < allsecond4)
                    {
                        zaotuitime++;
                    }
                    if(allsecond3 > allsecond5)
                    {
                        jiabantime = allsecond3 - allsecond5;
                    }
                    else if(allsecond3 < allsecond5)
                    {
                        jiabantime = 0;
                    }
                }

            }

        }

        //把表格里的数据转成数字，进行相加
        //转换缺勤
        QString time_str = ui->tableWidget_2->item(m,n)->text(); //时分秒

        QTime time=QTime::fromString(time_str);

        int allsecond = time.hour()*60*60+time.minute()*60;

        allsecond += queqintime;

        QString strtime;
        QString time_str1 = ui->tableWidget_2->item(m,n)->text().at(0);
        QString time_str2 = ui->tableWidget_2->item(m,n)->text().at(1);
        int time_num = (time_str1+time_str2).toInt();
        //qDebug()<<"24：："<<time_num;

        if(time_num == 24 || time_num == 16)
        {

            time_num += 8;

            strtime = QString::number(time_num) +":"+"00";

        }
        else
        {
            QTime a(0,0);
            a = a.addSecs(int(allsecond));

            strtime = a.toString("hh:mm");
        }


        //转换加班

        QString strtime2;
        QString jia_str = ui->tableWidget_2->item(m,n+1)->text(); //时分秒
        //qDebug()<<"jiajia"<<jia_str;

        QTime jia_time=QTime::fromString(jia_str);

        int jia_second = jia_time.hour()*60*60+jia_time.minute()*60;

        //qDebug()<<"secon"<<jia_second;

        jia_second += jiabantime;
        //qDebug()<<"alljia:"<<jia_second;
        //qDebug()<<"jiatime"<<jiabantime;

        QTime a(0,0);
        a = a.addSecs(int(jia_second));

        strtime2 = a.toString("hh:mm");



        //转换迟到
        int time_str3 = ui->tableWidget_2->item(m,n+2)->text().toInt();

        time_str3 += chidaotime;

        //转换早退
        int time_str4 = ui->tableWidget_2->item(m,n+3)->text().toInt();
        time_str4 += zaotuitime;

        ui->tableWidget_2->setItem(m, n, new QTableWidgetItem(strtime));
        ui->tableWidget_2->setItem(m, n+1, new QTableWidgetItem(strtime2));
        ui->tableWidget_2->setItem(m, n+2, new QTableWidgetItem(QString::number((time_str3))));
        ui->tableWidget_2->setItem(m, n+3, new QTableWidgetItem(QString::number((time_str4))));

        m++;

        queqintime = 0;

        jiabantime = 0;

        chidaotime = 0;

        zaotuitime = 0;

    }



}

