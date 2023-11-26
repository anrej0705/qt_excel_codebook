#include <QtCore/QCoreApplication>
#include <qsqldatabase.h>
#include <qsqlquery.h>
#include <qvariant.h>
#include <qdebug.h>
#include <QtWidgets>
#include <qsqlerror.h>

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);
	int rows = 0;
	int queryStart = 1;
	int queryLng = 0;

	int geraltExpressinsCount = 0;
	int trissExperrsionsCount = 0;
	int yenneferExpressionsCount = 0;
	int shaniExpressionsCount = 0;
	int keiraExpressionsCount = 0;
	int syannaExpressionsCount = 0;

	int hexIndexIn = 0;
	int charNameExpIndexIn = 0;

	std::string hexContainer = "";
	std::string hexContainer2 = "";

	QString desGeraltPath("P:\\w3utils_proper\\wav\\geraltExp\\");
	QString desTrissPath("P:\\w3utils_proper\\wav\\trissExp\\");
	QString desYenneferPath("P:\\w3utils_proper\\wav\\yenneferExp\\");
	QString desShaniPath("P:\\w3utils_proper\\wav\\shaniExp\\");
	QString desKeiraMetzPath("P:\\w3utils_proper\\wav\\keiraExp\\");
	QString desSyannaPath("P:\\w3utils_proper\\wav\\syannaExp\\");

	QFile srcName;
	QFile desName;
	std::string fileSuffix = ".wav.wav";
	QString srcPath("P:\\w3utils_proper\\w3speechConverted\\");
	QString srcAbsolutePath;
	QByteArray data;

	QRegExp geraltExp("\\s+(Geralt):");
	QRegExp trissExp("\\s+(Triss):");
	QRegExp yenneferExp("\\s+(Yennefer):");
	QRegExp shaniExp("\\s+(Shani):");
	QRegExp keiraExp("\\s+(Keira\\ Metz):");
	QRegExp syannaExp("\\s+(Syanna):");
	QRegExp hexExp("(0x[0-9a-fA-F]+)");
	QSqlError serr;
	QString err;
	QString querResult;
	QSqlDatabase db = QSqlDatabase::addDatabase("QODBC", "xlsx_connection");
	db.setDatabaseName("DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" + QString("C:\\cppt\\qt_excel_codebook\\qt_excel_codebook\\Witcher3DialogS.xlsx"));
	if (db.open())
	{
		queryStart = 1;
		queryLng = 32;
		for (int a = 0;a < 4800;++a)
		{
			QSqlQuery query("SELECT * FROM [" + QString("Dialog ID") + "$A" + QString::number(queryStart) + ":A" + QString::number(queryStart + queryLng) + "]", db);
			while (query.next())
			{
				QString column1 = QString(query.value(0).toString());
				//QString column1 = QString::fromStdString(std::string(QString(query.value(0).toString()).toStdString()).substr(0, 36));
				if (column1.contains(QRegExp("\\s*\\d+\\s+(0x[0-9a-fA-F]+\\s*Geralt:)\\s")))
				{
					//qDebug() << column1;
					hexIndexIn = hexExp.indexIn(column1);
					charNameExpIndexIn = geraltExp.indexIn(column1);
					hexContainer = std::string(column1.toStdString()).substr(hexIndexIn, 10);
					qDebug() << "HEX entry: " << QString::fromStdString(hexContainer) << " [Geralt] name entry: " << QString::number(charNameExpIndexIn);
					
					hexContainer2 = hexContainer;
					hexContainer += fileSuffix;
					
					srcAbsolutePath = srcPath + QString::fromStdString(hexContainer);
					//srcPath += QString::fromStdString(hexContainer);
					srcName.setFileName(srcAbsolutePath);
					if (srcName.open(QIODevice::ReadOnly))
						qDebug() << "Source opened";
					else
						qDebug() << "Source not opened";
					data = srcName.readAll();
					srcName.close();
					
					srcAbsolutePath=desGeraltPath + QString::fromStdString(hexContainer2) + "_geralt" + QString(".wav");
					
					desName.setFileName(srcAbsolutePath);
					if (desName.open(QIODevice::WriteOnly))
						qDebug() << "Destination file created";
					else
						qDebug() << "Destination file not created";
					desName.write(data);
					desName.close();

					
					//querResult.append(column1);
					//querResult.append("\n");
					++geraltExpressinsCount;
				}
				else if (column1.contains(QRegExp("\\s*\\d+\\s+(0x[0-9a-fA-F]+\\s*Triss:)\\s")))
				{
					//qDebug() << column1;
					hexIndexIn = hexExp.indexIn(column1);
					charNameExpIndexIn = trissExp.indexIn(column1);
					hexContainer = std::string(column1.toStdString()).substr(hexIndexIn, 10);
					qDebug() << "HEX entry: " << QString::fromStdString(hexContainer) << " [Triss] name entry: " << QString::number(charNameExpIndexIn);
					
					hexContainer2 = hexContainer;
					hexContainer += fileSuffix;
					
					srcAbsolutePath = srcPath + QString::fromStdString(hexContainer);
					//srcPath += QString::fromStdString(hexContainer);
					srcName.setFileName(srcAbsolutePath);
					if (srcName.open(QIODevice::ReadOnly))
						qDebug() << "Source opened";
					else
						qDebug() << "Source not opened";
					data = srcName.readAll();
					srcName.close();
					
					srcAbsolutePath=desTrissPath + QString::fromStdString(hexContainer2) + "_triss" + QString(".wav");
					
					desName.setFileName(srcAbsolutePath);
					if (desName.open(QIODevice::WriteOnly))
						qDebug() << "Destination file created";
					else
						qDebug() << "Destination file not created";
					desName.write(data);
					desName.close();

					//querResult.append(column1);
					//querResult.append("\n");
					++trissExperrsionsCount;
				}
				else if (column1.contains(QRegExp("\\s*\\d+\\s+(0x[0-9a-fA-F]+\\s*Yennefer:)\\s")))
				{
					//qDebug() << column1;
					hexIndexIn = hexExp.indexIn(column1);
					charNameExpIndexIn = yenneferExp.indexIn(column1);
					hexContainer = std::string(column1.toStdString()).substr(hexIndexIn, 10);
					qDebug() << "HEX entry: " << QString::fromStdString(hexContainer) << " [Yennefer] name entry: " << QString::number(charNameExpIndexIn);

					hexContainer2 = hexContainer;
					hexContainer += fileSuffix;

					srcAbsolutePath = srcPath + QString::fromStdString(hexContainer);
					//srcPath += QString::fromStdString(hexContainer);
					srcName.setFileName(srcAbsolutePath);
					if (srcName.open(QIODevice::ReadOnly))
						qDebug() << "Source opened";
					else
						qDebug() << "Source not opened";
					data = srcName.readAll();
					srcName.close();

					srcAbsolutePath = desYenneferPath + QString::fromStdString(hexContainer2) + "_yennefer" + QString(".wav");

					desName.setFileName(srcAbsolutePath);
					if (desName.open(QIODevice::WriteOnly))
						qDebug() << "Destination file created";
					else
						qDebug() << "Destination file not created";
					desName.write(data);
					desName.close();


					//querResult.append(column1);
					//querResult.append("\n");
					++yenneferExpressionsCount;
				}
				else if (column1.contains(QRegExp("\\s*\\d+\\s+(0x[0-9a-fA-F]+\\s*Shani:)\\s")))
				{
					//qDebug() << column1;
					hexIndexIn = hexExp.indexIn(column1);
					charNameExpIndexIn = shaniExp.indexIn(column1);
					hexContainer = std::string(column1.toStdString()).substr(hexIndexIn, 10);
					qDebug() << "HEX entry: " << QString::fromStdString(hexContainer) << " [Shani] name entry: " << QString::number(charNameExpIndexIn);

					hexContainer2 = hexContainer;
					hexContainer += fileSuffix;

					srcAbsolutePath = srcPath + QString::fromStdString(hexContainer);
					//srcPath += QString::fromStdString(hexContainer);
					srcName.setFileName(srcAbsolutePath);
					if (srcName.open(QIODevice::ReadOnly))
						qDebug() << "Source opened";
					else
						qDebug() << "Source not opened";
					data = srcName.readAll();
					srcName.close();

					srcAbsolutePath = desShaniPath + QString::fromStdString(hexContainer2) + "_shani" + QString(".wav");

					desName.setFileName(srcAbsolutePath);
					if (desName.open(QIODevice::WriteOnly))
						qDebug() << "Destination file created";
					else
						qDebug() << "Destination file not created";
					desName.write(data);
					desName.close();


					//querResult.append(column1);
					//querResult.append("\n");
					++shaniExpressionsCount;
				}
				else if (column1.contains(QRegExp("\\s*\\d+\\s+(0x[0-9a-fA-F]+\\s*Keira\\ Metz+:)\\s")))
				{
					//qDebug() << column1;
					hexIndexIn = hexExp.indexIn(column1);
					charNameExpIndexIn = keiraExp.indexIn(column1);
					hexContainer = std::string(column1.toStdString()).substr(hexIndexIn, 10);
					qDebug() << "HEX entry: " << QString::fromStdString(hexContainer) << " [Keira Metz] name entry: " << QString::number(charNameExpIndexIn);

					hexContainer2 = hexContainer;
					hexContainer += fileSuffix;

					srcAbsolutePath = srcPath + QString::fromStdString(hexContainer);
					//srcPath += QString::fromStdString(hexContainer);
					srcName.setFileName(srcAbsolutePath);
					if (srcName.open(QIODevice::ReadOnly))
						qDebug() << "Source opened";
					else
						qDebug() << "Source not opened";
					data = srcName.readAll();
					srcName.close();

					srcAbsolutePath = desKeiraMetzPath + QString::fromStdString(hexContainer2) + "_keira_metz" + QString(".wav");

					desName.setFileName(srcAbsolutePath);
					if (desName.open(QIODevice::WriteOnly))
						qDebug() << "Destination file created";
					else
						qDebug() << "Destination file not created";
					desName.write(data);
					desName.close();


					//querResult.append(column1);
					//querResult.append("\n");
					++keiraExpressionsCount;
				}
				else if (column1.contains(QRegExp("\\s*\\d+\\s+(0x[0-9a-fA-F]+\\s*Syanna:)\\s")))
				{
					//qDebug() << column1;
					hexIndexIn = hexExp.indexIn(column1);
					charNameExpIndexIn = syannaExp.indexIn(column1);
					hexContainer = std::string(column1.toStdString()).substr(hexIndexIn, 10);
					qDebug() << "HEX entry: " << QString::fromStdString(hexContainer) << " [Syanna] name entry: " << QString::number(charNameExpIndexIn);

					hexContainer2 = hexContainer;
					hexContainer += fileSuffix;

					srcAbsolutePath = srcPath + QString::fromStdString(hexContainer);
					//srcPath += QString::fromStdString(hexContainer);
					srcName.setFileName(srcAbsolutePath);
					if (srcName.open(QIODevice::ReadOnly))
						qDebug() << "Source opened";
					else
						qDebug() << "Source not opened";
					data = srcName.readAll();
					srcName.close();

					srcAbsolutePath = desSyannaPath + QString::fromStdString(hexContainer2) + "_syanna" + QString(".wav");

					desName.setFileName(srcAbsolutePath);
					if (desName.open(QIODevice::WriteOnly))
						qDebug() << "Destination file created";
					else
						qDebug() << "Destination file not created";
					desName.write(data);
					desName.close();


					//querResult.append(column1);
					//querResult.append("\n");
					++syannaExpressionsCount;
				}
			}
			queryStart += queryLng;
		}
		qDebug() << "Found expressions count:\n"
			<< " Geralt: " << geraltExpressinsCount << "Exp" << " approx lng: " << geraltExpressinsCount * 4 << " sec\n\n"
			<< " Triss: " << trissExperrsionsCount << "Exp" << " approx lng: " << trissExperrsionsCount * 4 << " sec\n"
			<< " Yennefer: " << yenneferExpressionsCount << "Exp" << " approx lng: " << yenneferExpressionsCount * 4 << " sec\n"
			<< " Shani: " << shaniExpressionsCount << "Exp" << " approx lng: " << shaniExpressionsCount * 4 << " sec\n"
			<< " Keira: " << keiraExpressionsCount << "Exp" << " approx lng: " << keiraExpressionsCount * 4 << " sec\n"
			<< " Syanna: " << syannaExpressionsCount << "Exp" << " approx lng: " << syannaExpressionsCount * 4 << " sec\n";
		db.close();
		QSqlDatabase::removeDatabase("xlsx_connection");
	}
	else
	{
		serr = db.lastError();
		err = serr.text();
	}
    return a.exec();
}
