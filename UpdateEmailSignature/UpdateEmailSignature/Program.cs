using System;
using System.Runtime.CompilerServices;

using static Const;
using static Functions;

Directory.CreateDirectory("./log");
Directory.CreateDirectory("./conf");

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
string website;
string b2c;
string vkLink;
string okLink;
string telegramLink;
string rutubeLink;
string сityStreet;



using (StreamReader ReaderObject = new StreamReader(configFilelway))
{
    string[] fileData = new string[] { };
    fileData = ReaderObject.ReadToEnd().Split('\n');
    sqlServerName = GetParametrValue(fileData[0], ": ")[1];
    sqlNameDB = GetParametrValue(fileData[1], ": ")[1];
    sqlUserName = GetParametrValue(fileData[2], ": ")[1];
    sqlPassword = GetParametrValue(fileData[3], ": ")[1];
    ldapPath = GetParametrValue(fileData[4], ": ")[1];
    adTagMailIgnor = GetParametrValue(fileData[5], ": ")[1];
    company = GetParametrValue(fileData[6], ": ")[1];
    streetAddress = GetParametrValue(fileData[7], ": ")[1];
    city = GetParametrValue(fileData[8], ": ")[1];
    telephoneNumber = GetParametrValue(fileData[9], ": ")[1];
    website = GetParametrValue(fileData[10], ": ")[1];
    b2c = GetParametrValue(fileData[11], ": ")[1];
    vkLink = GetParametrValue(fileData[12], ": ")[1];
    okLink = GetParametrValue(fileData[13], ": ")[1];
    telegramLink = GetParametrValue(fileData[14], ": ")[1];
    rutubeLink = GetParametrValue(fileData[15], ": ")[1];
}

List<string> longTitleNames = new List<string>();
List<string> shortTitleNames = new List<string>();
using (StreamReader ReaderObject = new StreamReader(titleListFileWay))
{
    string[] fileData = new string[] { };
    fileData = ReaderObject.ReadToEnd().Split('\n');
    for (int i = 0; i < fileData.Length; i++)
    {
        longTitleNames.Add(GetParametrValue(fileData[i], " :")[0]);
        shortTitleNames.Add(GetParametrValue(fileData[i], ": ")[1]);
    }
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

    if (userDataList[i]["company"] != companyName)
    {
        userDataList[i]["company"] = "&nbsp;";
    }
    if (userDataList[i]["streetAddress"] == null)
    {
        userDataList[i]["streetAddress"] = streetAddress;
    }
    if (userDataList[i]["l"] == null)
    {
        userDataList[i]["l"] = city.Substring(0, city.Length - 1);
    }
    if (userDataList[i]["mobile"] != null)
    {
        userDataList[i]["mobile"] = ", " + userDataList[i]["mobile"];
    }
    for (int j = 0; j < longTitleNames.Count; j++)
    {
        if (userDataList[i]["title"] == longTitleNames[j])
        {
            userDataList[i]["title"] = shortTitleNames[j];
        }
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
        сityStreet = userDataList[i]["l"] + ", " + userDataList[i]["streetAddress"];

        signatureText = "<div style=\"background-color: #ffffff; max-width: 800px; font-family: 'Montserrat', Arial, sans-serif; margin: 0 left; padding: 10px 0;\">" +
                    "\r\n<table style=\"background-image: url('https://mail.sstkvik.ru/signature/back.jpg'); border-spacing: 0px;\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\">" +
                    "\r\n<tbody>" +
                    "\r\n<tr>" +
                    "\r\n<td style=\"color: white; font-size: 14px; min-width: 350px; max-width: 300px;\">" +
                    "\r\n<div style=\"padding-left: 20px; max-width: 315px; padding-top: 5px;\">С уважением,</div>" +
                    "\r\n<div style=\"padding-left: 20px; font-weight: bold;\">" + userDataList[i]["cn"] + "</div>" +
                    "\r\n<hr style=\"border: 0; height: 2px; background-color: #8b00ff; width: 73%; margin-top: 2px; margin-bottom: 2px; margin-left: 20px;\" align=\"left\" />" +
                    "\r\n<div style=\"padding-left: 20px;\">" + userDataList[i]["title"] + "</div>" +
                    "\r\n<div style=\"padding-left: 20px; padding-bottom: 5px;\">" + userDataList[i]["company"] + "</div>" +
                    "\r\n</td>" +
                    "\r\n<td style=\"vertical-align: center; max-width: 365;\">" +
                    "\r\n<table style=\"border-spacing: 0px; width: 100%;\" cellspacing=\"0\" cellpadding=\"0\">" +
                    "\r\n<tbody>" +
                    "\r\n<tr>" +
                    "\r\n<td style=\"vertical-align: center; width: 70%;\">" +
                    "\r\n<table style=\"border-spacing: 1; margin-bottom: 0px;\" cellspacing=\"0\" cellpadding=\"0\">" +
                    "\r\n<tbody>" +
                    "\r\n<tr>" +
                    "\r\n<td style=\"padding-left: 20px; padding-bottom: 2px; vertical-align: middle;\" width=\"20\"><img src=\"https://mail.sstkvik.ru/signature/phone.png\" alt=\"Телефон\" width=\"25\" height=\"25\" /></td>" +
                    "\r\n<td style=\"padding-left: 8px; font-size: 12px;\">" + telephoneNumber + " (доб. " + userDataList[i]["telephoneNumber"] + ")" + userDataList[i]["mobile"] + "</td>" +
                    "\r\n</tr>" +
                    "\r\n</tbody>" +
                    "\r\n</table>" +
                    "\r\n<table style=\"border-spacing: 1; margin-bottom: 0px;\" cellspacing=\"0\" cellpadding=\"0\">" +
                    "\r\n<tbody>" +
                    "\r\n<tr>" +
                    "\r\n<td style=\"padding-left: 20px; padding-bottom: 2px; vertical-align: middle;\" width=\"20\"><img src=\"https://mail.sstkvik.ru/signature/addres.png\" alt=\"Адрес\" width=\"25\" height=\"25\" /></td>" +
                    "\r\n<td style=\"padding-left: 8px; font-size: 13px;\">" + сityStreet + "</td>" +
                    "\r\n</tr>" +
                    "\r\n</tbody>" +
                    "\r\n</table>" +
                    "\r\n<table style=\"border-spacing: 1; margin-bottom: 0px;\" cellspacing=\"0\" cellpadding=\"0\">" +
                    "\r\n<tbody>" +
                    "\r\n<tr>" +
                    "\r\n<td style=\"padding-left: 20px; vertical-align: middle;\" width=\"20\"><img src=\"https://mail.sstkvik.ru/signature/site.png\" alt=\"Сайт\" width=\"25\" height=\"25\" /></td>" +
                    "\r\n<td style=\"padding-left: 8px;\"><a style=\"color: #333; text-decoration: none; font-size: 13px;\" href=\"https://" + website + "\">" + website + "</a></td>" +
                    "\r\n</tr>" +
                    "\r\n</tbody>" +
                    "\r\n</table>" +
                    "\r\n</td>" +
                    "\r\n<td style=\"padding-left: 0px; vertical-align: center; text-align: center; width: 20%;\"><img src=\"https://mail.sstkvik.ru/signature/qr-code/QR-code.png\" alt=\"QR-код\" width=\"75\" height=\"75\" /></td>" +
                    "\r\n</tr>" +
                    "\r\n</tbody>" +
                    "\r\n</table>" +
                    "\r\n</td>" +
                    "\r\n</tr>" +
                    "\r\n</tbody>" +
                    "\r\n</table>" +
                    "\r\n<div style=\"margin-top: 20px;\"><img src=\"https://mail.sstkvik.ru/signature/car.gif\" width=\"800\" /></div>\r\n<table style=\"border-spacing: 0;\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\">" +
                    "\r\n<tbody>" +
                    "\r\n<tr>" +
                    "\r\n<td style=\"padding-top: 0px; vertical-align: top; width: 50%;\">" +
                    "\r\n<table style=\"border-spacing: 0;\" cellspacing=\"0\" cellpadding=\"0\">" +
                    "\r\n<tbody>" +
                    "\r\n<tr>" +
                    "\r\n<td style=\"padding-right: 20px; padding-left: 20px; padding-bottom: 5px;\"><a style=\"text-decoration: none;\" href=\"https://" + vkLink + "\"> <img src=\"https://mail.sstkvik.ru/signature/VK.png\" alt=\"ВК\" width=\"30\" height=\"30\" /> </a></td>" +
                    "\r\n<td style=\"padding-right: 20px;\"><a style=\"text-decoration: none; padding-bottom: 5px;\" href=\"https://" + telegramLink + "\"> <img src=\"https://mail.sstkvik.ru/signature/Telegram.png\" alt=\"ТЕЛЕГРАМ\" width=\"30\" height=\"30\" /> </a></td>" +
                    "\r\n<td style=\"padding-right: 20px;\"><a style=\"text-decoration: none; padding-bottom: 5px;\" href=\"https://" + rutubeLink + "\"> <img src=\"https://mail.sstkvik.ru/signature/Rutube.png\" alt=\"РУТУБ\" width=\"30\" height=\"30\" /> </a></td>" +
                    "\r\n<td style=\"padding-right: 20px;\"><a style=\"text-decoration: none; padding-bottom: 5px;\" href=\"https://" + okLink + "\"> <img src=\"https://mail.sstkvik.ru/signature/ok.png\" alt=\"ОДНОКЛАССНИКИ\" width=\"30\" height=\"30\" /> </a></td>" +
                    "\r\n</tr>" +
                    "\r\n</tbody>" +
                    "\r\n</table>" +
                    "\r\n</td>" +
                    "\r\n<td style=\"text-align: right; vertical-align: center; margin-bottom: 3px;\"><a href=\"https://yandex.ru/maps/org/61800352341\"><img src=\"https://mail.sstkvik.ru/signature/yandexStar.png\" alt=\"Яндекс\" width=\"297\" height=\"17\" /></a></td>" +
                    "\r\n</tr>" +
                    "\r\n</tbody>" +
                    "\r\n</table>" +
                    "\r\n<a href = \"https://sstkvik.ru/formail/link.php\" rel = \"noopener noreferrer\"><img src = \"https://sstkvik.ru/formail/index.php\" /></a></div>" +
                    "\r\n</div>";

        if (identity_id != "" && GetSqlSignature(connectToSQL, identity_id) != signatureText)
        {
            UsingStreamWriter("Сформированная для identity_id=" + identity_id +  " подпись не совпадает с полученной - перезаписываем.");
            UpdateSqLSignature(connectToSQL, signatureText, identity_id);
        }
    }
}

UsingStreamWriter("||===================Скрипт завершил работу===================||");