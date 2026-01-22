using System;
using System.Runtime.CompilerServices;

using static Const;
using static Functions;

Directory.CreateDirectory("./log");

string sqlServerName;
string sqlNameDB;
string sqlUserName;
string sqlPassword;
string ldapPath;
string adTagMailIgnor;
string company;
string streetAddress;
string city;
string telephoneNumber;



using (StreamReader ReaderObject = new StreamReader(configFilelway))
{
    string[] fileData = new string[] { };
    fileData = ReaderObject.ReadToEnd().Split('\n');
    sqlServerName = GetParametrValue(fileData[0], " ")[1];
    sqlNameDB = GetParametrValue(fileData[1], " ")[1];
    sqlUserName = GetParametrValue(fileData[2], " ")[1];
    sqlPassword = GetParametrValue(fileData[3], " ")[1];
    ldapPath = GetParametrValue(fileData[4], " ")[1];   
    adTagMailIgnor = GetParametrValue(fileData[5], " ")[1];
    company = GetParametrValue(fileData[6], ": ")[1];
    streetAddress = GetParametrValue(fileData[7], ": ")[1];
    city = GetParametrValue(fileData[8], ": ")[1];
    telephoneNumber = GetParametrValue(fileData[9], ": ")[1];
}

string connectToSQL =
$"Data Source={sqlServerName};" +
    $"Initial Catalog={sqlNameDB};" +
    $"User ID={sqlUserName};" +
    $"Password={sqlPassword};" +
    $"Encrypt=false;" +
    $"TrustServerCertificate=true";

Dictionary<string, string> userData = new Dictionary<string, string>
{
    {"cn", "" },
    {"title", "" },
    {"company", "" },
    {"streetAddress", "" },
    {"l", "" },
    {"mail", "" },
    {"telephoneNumber", "" },
    {"sAMAccountName", "" },
    {"mobile", "" },
    {"pager", "" }
};


UsingStreamWriter("||====================Значения поумолчанию====================||");
UsingStreamWriter("Компания: " + company);
UsingStreamWriter("Улица: " + streetAddress);
UsingStreamWriter("Город: " + city);
UsingStreamWriter("Телефон: " + telephoneNumber);

List<Dictionary<string, string>> userDataList = new List<Dictionary<string, string>>();
UsingStreamWriter("||============Получаем список пользователей из AD=============||");
List<string> adLoginList = new List<string>();
adLoginList = CutNullData(GetAllLoginAd(ldapPath).ToArray());


UsingStreamWriter("||=================Добавляем логины в словарь=================||");
for (int i = 0; i < adLoginList.Count; i++)
{
    Dictionary<string, string> userDataCopy = new Dictionary<string, string>(userData);
    userDataCopy["sAMAccountName"] = adLoginList[i];
    userDataList.Add(userDataCopy);
}


UsingStreamWriter("||=================Собираем данные по логинам=================||");
for (int i = 0; i < userDataList.Count; i++)
{
    Dictionary<string, string> userDataSearch = GetAdUserAtributs(ldapPath, userDataList[i]["sAMAccountName"]);

    userDataList[i]["cn"] = userDataSearch["cn"];
    userDataList[i]["title"] = userDataSearch["title"];
    userDataList[i]["company"] = userDataSearch["company"];
    userDataList[i]["streetAddress"] = userDataSearch["streetAddress"];
    userDataList[i]["mail"] = userDataSearch["mail"];
    userDataList[i]["telephoneNumber"] = userDataSearch["telephoneNumber"];
    userDataList[i]["mobile"] = userDataSearch["mobile"];
    userDataList[i]["l"] = userDataSearch["l"];
    userDataList[i]["pager"] = userDataSearch["pager"];

    if (userDataList[i]["company"] == null)
    {
        userDataList[i]["company"] = "ООО «Компания В.И.К»";
    }
    if (userDataList[i]["streetAddress"] == null)
    {
        userDataList[i]["streetAddress"] = "ул. Ростовское шоссе, 66";
    }
    if (userDataList[i]["l"] == null)
    {
        userDataList[i]["l"] = "г. Краснодар";
    }
    if (userDataList[i]["telephoneNumber"] == null)
    {
        userDataList[i]["telephoneNumber"] = telephoneNumber;
    }
}


UsingStreamWriter("||===============Удаляем записи с пустыми mail================||");
userDataList.RemoveAll(dict => (dict["mail"] == null));

UsingStreamWriter("||=============Email, которые следует пропустить==============||");
for (int i = 0;  i < userDataList.Count; i++)
{
    if (userDataList[i]["pager"] == adTagMailIgnor)
    {
        UsingStreamWriter("cn=" + userDataList[i]["cn"]);
        UsingStreamWriter("mail=" + userDataList[i]["mail"]);
        UsingStreamWriter("");
    }
}


UsingStreamWriter("||===============Удаляем записи с pager=" + adTagMailIgnor + "================||");
userDataList.RemoveAll(dict => (dict["pager"] == adTagMailIgnor));


UsingStreamWriter("||====Выносим юзеров с одинаковыми mail в отдельный список====||");
UsingStreamWriter("");
var allDuplicates = userDataList
    .GroupBy(dict => dict["mail"])
    .Where(group => group.Count() > 1)
    .SelectMany(group => group)
    .ToList();
for (int i = 0;  i < allDuplicates.Count; i++)
{
    UsingStreamWriter("cn=" + allDuplicates[i]["cn"]);
    UsingStreamWriter("mail=" + allDuplicates[i]["mail"]);
    UsingStreamWriter("");
}


UsingStreamWriter("||===Удаляем из основного списка записи с задвоенными mail====||");
for (int i = 0;  i < allDuplicates.Count; i++)
{
    userDataList.RemoveAll(dict => (dict["mail"] == allDuplicates[i]["mail"]));
}

UsingStreamWriter("||==========Сравниваем наборы mail из AD с RoundCube==========||");
UsingStreamWriter("||===============Список mail, которых нет в AD================||");
List<string> mailFromSql = GetSqlAllMail(connectToSQL);
for (int i = 0; i < mailFromSql.Count; i++)
{
    bool x = false;
    for (int j = 0; j < userDataList.Count; j++)
    {
        if (mailFromSql[i] == userDataList[j]["mail"])
        {
            x = true;
        }
    }
    if (x == false)
    {
        UsingStreamWriter(mailFromSql[i]);
    }
}


UsingStreamWriter("||===========У следующих mail не совпадает ФИО с AD===========||");
for (int i = 0; i < userDataList.Count; i++)
{
    FindeNameDoesntExistAd(connectToSQL, userDataList[i]["mail"], userDataList[i]["cn"]);
}


UsingStreamWriter("||=============Получаем identity_id по ФИО и mail=============||");
string signatureText = "";
for (int i = 0; i < userDataList.Count; i++)
{
    string identity_id = GetSqlIdentityID(connectToSQL, userDataList[i]["cn"], userDataList[i]["mail"]);
    if (identity_id != null)
    {
        if (userDataList[i]["mobile"] == null)
        {
            signatureText = openDiv + openSpan + "С уважением," + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["cn"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["title"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["company"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["l"] + ", " + userDataList[i]["streetAddress"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + "Mail:         " + closeSpan +
                openSpan + openStrong + "<a href=\"mailto:" + userDataList[i]["mail"] + "\">" + userDataList[i]["mail"] + "</a>" + closeStrong + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + "Phone:      8 (800)-222-52-62 доп:" + userDataList[i]["telephoneNumber"] + "  " + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + openStrong + "website     " + "<a href=\"https://" + website + "\" rel=\"noopener\">" + website + "</a>" + closeStrong + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + openStrong + "b2c          " + "<a href=\"https://" + b2c + "\" rel=\"noopener\">" + b2c + "</a>" + closeStrong + closeSpan + closeDiv + "\r\n" +
                openDiv + closeDiv + "\r\n" + openDiv + "<a href=\"https://" + websiteLink + "\" rel=\"noopener noreferrer\">" + "<img src=\"https://" + imgLink + "\" />" + "</a>" + closeDiv;
        }
        else
        {
            signatureText = openDiv + openSpan + "С уважением," + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["cn"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["title"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["company"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + userDataList[i]["l"] + ", " + userDataList[i]["streetAddress"] + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + "Mail:         " + closeSpan +
                openSpan + openStrong + "<a href=\"mailto:" + userDataList[i]["mail"] + "\">" + userDataList[i]["mail"] + "</a>" + closeStrong + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + "Phone:      8 (800)-222-52-62 доп:" + userDataList[i]["telephoneNumber"] + "  " + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + "Mobile:      " + userDataList[i]["mobile"] + "  " + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + openStrong + "website     " + "<a href=\"https://" + website + "\" rel=\"noopener\">" + website + "</a>" + closeStrong + closeSpan + closeDiv + "\r\n" +
                openDiv + openSpan + openStrong + "b2c          " + "<a href=\"https://" + b2c + "\" rel=\"noopener\">" + b2c + "</a>" + closeStrong + closeSpan + closeDiv + "\r\n" +
                openDiv + closeDiv + "\r\n" + openDiv + "<a href=\"https://" + websiteLink + "\" rel=\"noopener noreferrer\">" + "<img src=\"https://" + imgLink + "\" />" + "</a>" + closeDiv;
        }
            
        if (identity_id != "" && GetSqlSignature(connectToSQL, identity_id) != signatureText)
        {
            UsingStreamWriter("Сформированная для identity_id=" + identity_id +  " подпись не совпадает с полученной - перезаписываем.");
            UpdateSqLSignature(connectToSQL, signatureText, identity_id);
        }
    }
}

UsingStreamWriter("||===================Скрипт завершил работу===================||");