#include "frmmain.h"
#include "ui_frmmain.h"
#include <QFileDialog>
#include <QMessageBox>
#include "sendemailapi/smtpmime.h"
#include "kexcelreader.h"
#include <QDebug>
#include <QAxObject>
#include <QSet>
frmMain::frmMain(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::frmMain)
{
    ui->setupUi(this);
}

frmMain::~frmMain()
{
    delete ui;
}

void frmMain::on_btnSelect_clicked()
{
    QFileDialog dialog(this);
    dialog.setFileMode(QFileDialog::ExistingFiles);

    if (dialog.exec()){
        ui->txtAtta->clear();
        QStringList files=dialog.selectedFiles();
        foreach (QString file, files) {
            ui->txtAtta->setText(ui->txtAtta->text()+file);
        }
    }
}

bool frmMain::Check()
{
    if (ui->txtSender->text()==""){
        QMessageBox::critical(this,"错误","用户名不能为空!");
        ui->txtSender->setFocus();
        return false;
    }
    if (ui->txtSenderPwd->text()==""){
        QMessageBox::critical(this,"错误","用户密码不能为空!");
        ui->txtSenderPwd->setFocus();
        return false;
    }
    if (ui->txtAtta->text()==""){
        QMessageBox::critical(this,"错误","附件不能为空!");
        ui->txtAtta->setFocus();
        return false;
    }
    return true;
}

QMap<QString,QString> frmMain::ReadExcel(const QString& xlsFile){
    QMap<QString,QString> map;
    KExcelReader reader;
     if(reader.isExcelApplicationAvailable()){
       if(reader.open(xlsFile)){
           int sheetCount=reader.sheetCount();
           qDebug("sheetCount:%d",sheetCount);

           QAxObject* workbook=reader.getWorkbook();
           for(int i=1; i<=sheetCount; i++)
           {

           int rowCount=100;
           int columnCount=30;
           QList<QVariantList> list=reader.values(columnCount,rowCount,i);
           qDebug("size:%d",  list.size());


           map=logicHtml(list);

           }


       }else{
          QMessageBox::critical(this,"错误","无法打开Excel");
       }

     }else{
          QMessageBox::critical(this,"错误","Excel不支持");
     }
    return map;
}

QMap<QString,QString> frmMain::logicHtml(QList<QVariantList>& list){

      QMap<QString,QString> map;
      QVariantList titleList;
      QSet<QString> emailSet;

      for(int j=0;j<list.size();j++){
          if(j==0) {
              titleList=list.at(j);
              list.removeAt(j);
              continue;
          }

          QVariantList tmpList=list.at(j);
          QVariant temp=tmpList.at(tmpList.size()-1);

          if(temp.isNull()) continue;
          emailSet.insert(temp.toString());
      }

      QSetIterator<QString> i(emailSet);
      while (i.hasNext()) {
          QString email=i.next();
          //qDebug() << i.next();
          QString html="<html>";
          html+="<head>";
          html+=" <title></title>";
          html+="</head>";
          html+="<body>";
          html+="<div style=\"font-size: 14px; line-height: 21px;\">";
          html+="<span style=\"font-family: 宋体; font-size: 27px; line-height: 48px;\">Dear，</span>";
          html+="</div>";
          html+="<div style=\"font-size: 14px; line-height: 21px;\">";
          html+="<span style=\"font-family: 宋体; font-size: 27px; line-height: 48px;\">下面是您"+ui->month->currentText()+"份实发的报销明细，请查收！如有疑问请与我联系！</span>";
          html+="</div>";

          html+="<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"1300\" style=\"border-collapse:collapse;\">";
          html+="<tbody>";
          html+="<tr height=\"50\" style=\"mso-height-source:userset;height:37.2pt\">";

          //遍历邮件title
          for(int j=0;j<titleList.size();j++){
            QVariant var=titleList.at(j);

            html+="<td height=\"50\" class=\"xl94\" width=\"10%\" style=\"height: 17.2pt;  padding: 0px; color: windowtext; font-size: 9pt;";
            html+="font-family: 微软雅黑, sans-serif; vertical-align: middle; border: 0.5pt solid windowtext; text-align: center; background-color: rgb(53, 164, 67);\">";
            html+=var.toString();
            html+="</td>";
          }
          html+="</tr>";

          //遍历内容
          for(int j=0;j<list.size();j++){
             QVariantList tmpList=list.at(j);
             QVariant emailStr=tmpList.at(tmpList.size()-1);
             if(email!=emailStr.toString()) continue;//判断收件人是否为同一个

             html+="<tr height=\"20\" style=\"mso-height-source:userset;height:15.0pt\">";
             for(int x=0;x<tmpList.size()-1;x++){//去掉excel里最后一列
                 QVariant tmp=tmpList.at(x);

                 html+="<td height=\"20\" class=\"xl101\" style=\"height: 15pt; border: 0.5pt solid windowtext; padding: 0px; color: windowtext; font-size: 8pt;";
                 html+=" font-family: 宋体; vertical-align: middle; white-space: nowrap;\">";
                 html+=tmp.toString();
                 html+="</td>";
             }
             html+="</tr>";
          }

          html+="</tbody>";
          html+="</table>";

          html+="</body>";
          html+="</html>";

          map.insert(email,html);
      }


   return map;

}


void frmMain::on_btnSend_clicked()
{
    if (!Check()){return;}

    //实例化发送邮件对象
    SmtpClient smtp(ui->cboxServer->currentText(),
                    ui->cboxPort->currentText().toInt(),
                    ui->ckSSL->isChecked()?SmtpClient::SslConnection:SmtpClient::TcpConnection);
    smtp.setUser(ui->txtSender->text());
    smtp.setPassword(ui->txtSenderPwd->text());


    if (!smtp.connectToHost()){
        QMessageBox::critical(this,"错误","服务器连接失败!");
        return;
    }
    if (!smtp.login()){
        QMessageBox::critical(this,"错误","用户登录失败!");
        return;
    }

    QMap<QString,QString> map=ReadExcel(ui->txtAtta->text());


    QMapIterator<QString, QString> m(map);
    while (m.hasNext()) {
         //  qDebug() <<m.next().key();
         //  qDebug() << m.value();


    //构建邮件主题,包含发件人收件人附件等.
    MimeMessage message;
    message.setSender(new EmailAddress(ui->txtSender->text()));

    //逐个添加收件人
    message.addRecipient(new EmailAddress(m.next().key()));

    //构建邮件标题
    message.setSubject("报销单");

    //构建邮件正文
    MimeHtml text;
    text.setHtml(m.value());
    message.addPart(&text);

    if (!smtp.sendMail(message)){
        QMessageBox::critical(this,"错误","邮件发送失败!");
        return;
    }
   }
    QMessageBox::information(this,"错误","邮件发送成功!");
    smtp.quit();
}

void frmMain::on_cboxServer_currentIndexChanged(int index)
{
    if (index==2){
        ui->cboxPort->setCurrentIndex(1);
        ui->ckSSL->setChecked(true);
    }else{
        ui->cboxPort->setCurrentIndex(0);
        ui->ckSSL->setChecked(false);
    }
}
