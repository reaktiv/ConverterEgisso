unit U_main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Buttons,
  Vcl.ExtCtrls, ComObj, Vcl.Grids, DateUtils, Vcl.Imaging.pngimage;

type
  TF_main = class(TForm)
    P_Tools: TPanel;
    B_OpenFile: TBitBtn;
    P_Button: TPanel;
    P_Log: TPanel;
    RE_Log: TRichEdit;
    OD_OpenXlsxFile: TOpenDialog;
    M_Info: TMemo;
    M_Error: TMemo;
    SG_Fact: TStringGrid;
    B_log: TBitBtn;
    B_Grid: TBitBtn;
    M_CSV: TMemo;
    B_Info: TBitBtn;
    P_Info: TPanel;
    Label1: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label8: TLabel;
    Image1: TImage;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    L_VersionSoftware: TLabel;
    L_VersionTemplate: TLabel;
    L_DateBuild: TLabel;

    function DeleteEverExNumber(Text:String) : String;
    function ReplaceSexMark(Text:String): String;
    function ReplaceNumberSeparator(Text:String) : String;
    function IsDate(Text:String): Boolean;
    function IsFloat(Text:String):Boolean;


    procedure getData(xlsxFilePath:string);
    procedure processingOrgInfo();
    procedure processingMassMszFact();

    procedure renderTableMszFact(); //Процедура отображает массив фактов назначения МСЗ
    procedure renderDataToCsv(); //Преобразует все данные из massMszFact в массив CSV


    procedure B_OpenFileClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure B_logClick(Sender: TObject);
    procedure B_GridClick(Sender: TObject);
    procedure B_InfoClick(Sender: TObject);



  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_main: TF_main;

  //Массивы
  massMszList : array of array of String; //Массив хранящий список доступных мер соц. поддержки
  massMszFact : array of array of String; //Массив хранящий список фактов назначения мер соц. поддержки
  massOrgList : array of array of String; //Массив хранящий список организаций
  orgInfo : array of String; //Массив хранящий данные учреждения

  XlsxFile : String; //Полный путь (с именем) к открываемому файлу-шаблону
  XlsxFileDir : String; //Путь к папке с открываемым файлом-шаблоном
  XlsxFileName : String; //Название файла с расширением (типом файла)
  XlsxFileNameNoExt : String; //Название файла без расширения

  M_Info : TMemo; //Хранит список не критических ошибок
  M_Error : TMemo; //Хранит список критических ошибок

const

  //Основные переменные
  VersionSoftware = '2.0'; //Версия программы
  VersionTemplate = '2.0'; //Версия шаблона с которым может работать программа
  DateBuild = '19.01.2021'; //Дата сборки

  //Переменные массива massOrgList
  massOrgListKratName = 0;
  massOrgListONMSZCode = 1;

  //Переменные массива massMszList
  massMszListKratName = 0;
  massMszListLMSZID = 1;
  massMszListСategoryID = 2;
  massMszListSumm = 3;

  //Переменные массива massMSZFact
  massMSZFactSNILS = 0;
  massMSZFactFamily = 1;
  massMSZFactName = 2;
  massMSZFactSurname = 3;
  massMSZFactGender = 4;
  massMSZFactBirthDate = 5;
  massMSZFactDecisionDate = 6;
  massMSZFactDateStart = 7;
  massMSZFactDateFinish = 8;
  massMSZFactLMSZName = 9;
  massMSZFactAmount = 10;
  massMSZFactLMSZID = 11; //Будет хранить LMSZID меры
  massMSZFactLMSZCategoryID= 12; //Будет хранить СategoryID меры

  //Переменные массива orgInfo
  orgInfoName = 0;
  orgInfoYear = 1;
  orgInfoMonth = 2;
  orgInfoFio = 3;
  orgInfoPhone = 4;
  orgInfoSpecialMarks = 5;
  orgInfoTemplateVersion = 6;
  orgInfoONMSZCode = 7; //Будет хранить код выбранной организации

  //Отступ в логах
  textIdent = '     '; //Хранит отступ в пробелах

implementation

{$R *.dfm}

//функция удаляет из переданной строки всё кроме цифр
function TF_main.DeleteEverExNumber(Text:String) : String;
  var i :integer;
  TextOut : String;
Begin
  TextOut := '';
  for i := 1 to length(Text) do
    if Text[i] in ['0'..'9'] then TextOut := TextOut + Text[i];
  result := TextOut;
End;


//Функция заменяет все "м" на "М" и все "ж" на "Ж"
function TF_main.ReplaceSexMark(Text:String): String;
Begin
  Text := StringReplace(Text, 'ж', 'Ж', [rfReplaceAll]);
  Text := StringReplace(Text, 'м', 'М', [rfReplaceAll]);
  Result := Text;
End;


//Функция заменяет все "." на ","
function TF_main.ReplaceNumberSeparator(Text:String) : String;
Begin
  Text := StringReplace(Text, '.', ',', [rfReplaceAll]);
  Result := Text;
End;


//Функция проверяет является ли переданная строка датой
function TF_main.IsDate(Text:String): Boolean;
  var dat:TDateTime;
Begin
  if TryStrToDate(Text,dat) then Result := true else Result := false;
End;


//Функция проверяет является ли переданная строка типом float
function TF_main.IsFloat(Text:String):Boolean;
  Var Mu:double;
begin
  if TryStrToFloat(Text, Mu) then Result := true else Result := false;
End;


//Процедура засовывает в таблицу массив фактов назначения мсз
procedure TF_Main.renderTableMszFact();
  var i,j : integer;
Begin
  //Обрабатываем массив
  if (Length(massMszFact) > 0)  then
    begin

      //Подготавливаем Грид
      SG_Fact.ColCount :=  14;
      SG_Fact.RowCount :=  Length(massMszFact) + 1;

      SG_Fact.FixedCols := 1;
      SG_Fact.FixedRows := 1;

      //Назначаем ширины и заголовки столбцам
        //Столбец номеров
        SG_Fact.ColWidths[0] := 60;
        SG_Fact.Cells[0,0] := '№ п.п';
        //Столбец СНИЛС
        SG_Fact.ColWidths[1] := 130;
        SG_Fact.Cells[1,0] := 'Снилс';
        //Столбец Фамилия
        SG_Fact.ColWidths[2] := 150;
        SG_Fact.Cells[2,0] := 'Фамилия';
        //Столбец Имя
        SG_Fact.ColWidths[3] := 150;
        SG_Fact.Cells[3,0] := 'Имя';
        //Столбец Отчество
        SG_Fact.ColWidths[4] := 150;
        SG_Fact.Cells[4,0] := 'Отчество';
        //Столбец Пол
        SG_Fact.ColWidths[5] := 50;
        SG_Fact.Cells[5,0] := 'Пол';
        //Столбец Дата рождения
        SG_Fact.ColWidths[6] := 120;
        SG_Fact.Cells[6,0] := 'Дата рождения';
        //Столбец Дата решения о назначении
        SG_Fact.ColWidths[7] := 120;
        SG_Fact.Cells[7,0] := 'Дата решения';
        //Столбец Дата начала действия
        SG_Fact.ColWidths[8] := 120;
        SG_Fact.Cells[8,0] := 'Дата начала';
        //Столбец Дата окончания действия
        SG_Fact.ColWidths[9] := 120;
        SG_Fact.Cells[9,0] := 'Дата окончания';
        //Столбец Мера
        SG_Fact.ColWidths[10] := 200;
        SG_Fact.Cells[10,0] := 'Мера';
        //Столбец Сумма
        SG_Fact.ColWidths[11] := 100;
        SG_Fact.Cells[11,0] := 'Сумма';
        //Столбец Идентификатор ЛМСЗ
        SG_Fact.ColWidths[12] := 300;
        SG_Fact.Cells[12,0] := 'Идентификатор ЛМСЗ';
        //Столбец Идентификатор ЛМСЗ
        SG_Fact.ColWidths[13] := 300;
        SG_Fact.Cells[13,0] := 'Идентификатор локальной категории';


      for i := Low(massMszFact) to High(massMszFact) do
         Begin
           SG_Fact.Cells[0,i+1] := IntToStr(i);
           SG_Fact.Cells[1,i+1] := massMszFact[i,massMSZFactSNILS];
           SG_Fact.Cells[2,i+1] := massMszFact[i,massMSZFactFamily];
           SG_Fact.Cells[3,i+1] := massMszFact[i,massMSZFactName];
           SG_Fact.Cells[4,i+1] := massMszFact[i,massMSZFactSurname];
           SG_Fact.Cells[5,i+1] := massMszFact[i,massMSZFactGender];
           SG_Fact.Cells[6,i+1] := massMszFact[i,massMSZFactBirthDate];
           SG_Fact.Cells[7,i+1] := massMszFact[i,massMSZFactDecisionDate];
           SG_Fact.Cells[8,i+1] := massMszFact[i,massMSZFactDateStart];
           SG_Fact.Cells[9,i+1] := massMszFact[i,massMSZFactDateFinish];
           SG_Fact.Cells[10,i+1] := massMszFact[i,massMSZFactLMSZName];
           SG_Fact.Cells[11,i+1] := massMszFact[i,massMSZFactAmount];
           SG_Fact.Cells[12,i+1] := massMszFact[i,massMSZFactLMSZID];
           SG_Fact.Cells[13,i+1] := massMszFact[i,massMSZFactLMSZCategoryID];
         End;
    end;

end;


//Процедура преобразует данные из massMSZFact в формат CSV
procedure TF_Main.renderDataToCsv();
  var i:Integer;
      delimiter : String;
begin
  //Подготавливаем Memo
  M_CSV.Clear;
  delimiter := ';';

  //Пишем заголовок в ЛОГ
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('Экспорт файлов');
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');

  //Записываем первую строку
  M_CSV.Lines.Add('RecType;LMSZID;categoryID;ONMSZCode;SNILS_recip;FamilyName_recip;Name_recip;' +
  'Patronymic_recip;Gender_recip;BirthDate_recip;doctype_recip;doc_Series_recip;doc_Number_recip;' +
  'doc_IssueDate_recip;doc_Issuer_recip;SNILS_reason;FamilyName_reason;Name_reason;Patronymic_reason;' +
  'Gender_reason;BirthDate_reason;doctype_reason;doc_Series_reason;doc_Number_reason;doc_IssueDate_reason;' +
  'doc_Issuer_reason;decision_date;dateStart;dateFinish;usingSign;criteria;FormCode;amount;measuryCode;' +
  'monetization;content;comment;equivalentAmount');

  //Накапливаем строки с данными
  for i := Low(massMszFact) to High(massMszFact) do
    begin
      M_CSV.Lines.Add(
      'Fact' + delimiter +
      massMszFact[i, massMSZFactLMSZID] + delimiter +
      massMszFact[i, massMSZFactLMSZCategoryID] + delimiter +
      orgInfo[orgInfoONMSZCode] + delimiter +
      massMszFact[i, massMSZFactSNILS] + delimiter +
      massMszFact[i, massMSZFactFamily] + delimiter +
      massMszFact[i, massMSZFactName] + delimiter +
      massMszFact[i, massMSZFactSurname] + delimiter +

      massMszFact[i, massMSZFactGender] + delimiter +
      massMszFact[i, massMSZFactBirthDate] +
      delimiter + delimiter + delimiter + delimiter + delimiter +   //Прибавляем 17 точек с запятой
      delimiter + delimiter + delimiter + delimiter + delimiter +
      delimiter + delimiter + delimiter + delimiter + delimiter +
      delimiter + delimiter +
      massMszFact[i, massMSZFactDecisionDate] + delimiter +
      massMszFact[i, massMSZFactDateStart] + delimiter +
      massMszFact[i, massMSZFactDateFinish] + delimiter +
      'Нет' + delimiter + delimiter + '01' + delimiter +
      massMszFact[i, massMSZFactAmount] + delimiter +
      '1' + delimiter + '0' + delimiter + delimiter + delimiter);
    end;

  //Выгружаем данные в csv файл
  if (DirectoryExists(XlsxFileDir + '\csv') <> true) then
    ForceDirectories(XlsxFileDir + '\csv'); //Если нужной папки нет - создаём.

  if (DirectoryExists(XlsxFileDir + '\csv') <> true) then
    M_Error.Lines.Add('Не удаётся создать папку ' + XlsxFileDir + '\csv ;'); //Пишем лог что не получилось создать папку

  if (M_Error.Text = '') then //Если на предыдущих этапах не было критических ошибок
    Begin
      //Удаляем файл если он существует
      if (FileExists(XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv') = true) then
        DeleteFile((XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv'));

      if (FileExists(XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv') <> true) then
        Begin
          M_CSV.Lines.SaveToFile(XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv');
          RE_Log.Lines.Add(textIdent + 'Данные выгружены в файл ' + XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv; ');
        end
      else
        M_Error.Lines.Add('Не удаётся создать файл ' + XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv'); //Пишем лог что не получилось создать файл
    End;



  //Выгружаем логи в файлы-логи
  if (DirectoryExists(XlsxFileDir + '\log') <> true) then
    ForceDirectories(XlsxFileDir + '\log'); //Если нужной папки нет - создаём.

  if (DirectoryExists(XlsxFileDir + '\log') <> true) then
    M_Error.Lines.Add('Не удаётся создать папку ' + XlsxFileDir + '\log ;'); //Пишем лог что не получилось создать папку

  if (M_Error.Text = '') then //Если на предыдущих этапах не было критических ошибок
    Begin
      //Удаляем файл если он существует
      if (FileExists(XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log') = true) then
        DeleteFile((XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log'));

      if (FileExists(XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log') <> true) then
        Begin
          RE_Log.Lines.Add(textIdent + 'Логи выгружены в файл ' + XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log; ');
          RE_Log.Lines.SaveToFile(XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log');
        end
      else
        M_Error.Lines.Add('Не удаётся создать файл ' + XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log'); //Пишем лог что не получилось создать файл
    End;

end;


//Процедура загружает данные в массивы из указанного файла
procedure TF_main.getData(xlsxFilePath:String);
  const
    xlCellTypeLastCell = $0000000B;
  var
    //Переменые для работы с файлом xlsx
    ExcelApp, ExcelSheet: OLEVariant;
    MyMass: Variant;
    i, x, y: Integer;

    BadLinesMsz : String; //Хранит номера некорретно заполненных строк справочника мер
    BadLinesOrg : String; //Хранит номера некорретно заполненных строк справочника учреждений
    BadLinesFact : String; //Хранит номера некорретно заполненных строк реестра фактов назначения МСЗ

Begin
  //Подготовка переменных
  BadLinesMsz := '';
  BadLinesOrg := '';
  BadLinesFact := '';


  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('Загрузка данных из файла ' + xlsxFilePath);
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');



  //создание OLE-объекта Excel
  RE_Log.Lines.Add(textIdent + 'Открытие файла...');
  ExcelApp := CreateOleObject('Excel.Application');
  //открытие книги Excel
  RE_Log.Lines.Add(textIdent + 'Открытие книги...');
  ExcelApp.Workbooks.Open(xlsxFilePath);


  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + 'Чтение справочника организаций');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // открытие листа книги
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Спр.Организаций'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Спр.Организаций'];

      // выделение последней задействованной ячейки на листе
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // получение значений размера выбранного диапазона
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // присвоение массиву диапазона ячеек на листе
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //Переносим данные в новый массив
      SetLength(massOrgList, 0); //Задаём нулевую размерность массива
      for i := 2 to x do  //Начиная со второй строки перебираем всё что осталось
        begin
          if ((VarToStr(MyMass[i, 1])<> '') and (VarToStr(MyMass[i, 2])<> '') and (VarToStr(MyMass[i, 3])<> '')) then  //Если первый и второй и третий столбцы заполнены
            Begin
              SetLength(massOrgList,Length(massOrgList) + 1); //Добавляем ещё одну строку
              SetLength(massOrgList[High(massOrgList)], 2); //В добавленной строке (Последней по счёту) задаём размерность 2 столбца
              massOrgList[High(massOrgList), massOrgListKratName] := MyMass[i, 2]; //Записываем название
              massOrgList[High(massOrgList), massOrgListONMSZCode] := MyMass[i, 3];   //Записываем код LMSZID
            End
          else
            Begin
              BadLinesOrg := BadLinesOrg + IntToStr(i) + ', ';
            End;

        end;

      //Выводим в логи количество найденных учреждений
      RE_Log.Lines.Add(textIdent + 'Всего найдено учреждений: ' + IntToStr(Length(massOrgList)));

      //Выводим список учреждений
      for i := Low(massOrgList) to High(massOrgList) do
        Begin
          RE_Log.Lines.Add(textIdent + IntToStr(i+1) + ')   ' + massOrgList[i,massOrgListKratName] + '   (ONMSZCode: ' +  massOrgList[i,massOrgListONMSZCode] + ');');
        end;

      //Заполняем список ошибок
      if (BadLinesOrg <> '') then
          M_Info.Lines.Add('Справочник организаций содержит некорректные строки (' + BadLinesOrg + '); ');

  RE_Log.Lines.Add('');



  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + 'Чтение справочника мер');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // открытие листа книги
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Спр.Меры'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Спр.Меры'];

      // выделение последней задействованной ячейки на листе
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // получение значений размера выбранного диапазона
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // присвоение массиву диапазона ячеек на листе
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //Перерабатываем найденные данные в нормальный массив
      SetLength(massMszList,0); //Задаём нулевую размерность массива
      for i := 2 to x do  //Начиная со второй строки перебираем всё что осталось
        begin
          if ((VarToStr(MyMass[i, 1])<> '') and (VarToStr(MyMass[i, 2])<> '') and (VarToStr(MyMass[i, 3])<> '') and (VarToStr(MyMass[i, 4])<> '')) then
            begin
              SetLength(massMszList,Length(massMszList) + 1); //Добавляем ещё одну строку
              SetLength(massMszList[High(massMszList)], 4); //В добавленной строке (Последней по счёту) задаём размерность 4 столбца
              massMszList[High(massMszList), massMszListKratName] := MyMass[i, 2]; //Записываем название
              massMszList[High(massMszList), massMszListLmszID] := MyMass[i, 3];   //Записываем код LMSZID
              massMszList[High(massMszList), massMszListСategoryID] := MyMass[i, 4];    //Записываем код categoryID
              massMszList[High(massMszList), massMszListSumm] := '0'; //Записываем ноль чтобы потом корректно приобразовать его в число и накапливать сумму
            end
          else
            Begin
              BadLinesMsz := BadLinesMsz + IntToStr(i) + ', ';
            End;
        end;

      //Выводим в логи количество найденных мер
      RE_Log.Lines.Add(textIdent + 'Всего найдено мер: ' + IntToStr(Length(massMszList)));

      //Выводим в логи список найденных мер
      for i := Low(massMszList) to High(massMszList) do
        Begin
          RE_Log.Lines.Add(textIdent + IntToStr(i+1) + ')   '+  massMszList[i,massMszListKratName] + '   (LMSZID:' +  massMszList[i,massMszListLmszID] + ', categoryID:' + massMszList[i,massMszListСategoryID] + ');');
        end;

      //Заполняем список ошибок
      if (BadLinesMsz <> '') then
          M_Info.Lines.Add('Справочник мер содержит некорректные строки (' + BadLinesMsz + '); ');

  RE_Log.Lines.Add('');



  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + 'Чтение реестра фактов назначения МСЗ');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // открытие листа книги
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Реестр фактов назначения МСЗ'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Реестр фактов назначения МСЗ'];

      // выделение последней задействованной ячейки на листе
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // получение значений размера выбранного диапазона
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // присвоение массиву диапазона ячеек на листе
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //Перерабатываем найденные данные в нормальный массив
      SetLength(massMszFact,0); //Задаём нулевую размерность массива
      for i := 2 to x do  //Начиная со второй строки перебираем всё что осталось
        begin
          if (
            (VarToStr(MyMass[i, 1])<> '') and (VarToStr(MyMass[i, 2])<> '') and (VarToStr(MyMass[i, 3])<> '') and
            (VarToStr(MyMass[i, 4])<> '') and (VarToStr(MyMass[i, 5])<> '') and (VarToStr(MyMass[i, 6])<> '') and
            (VarToStr(MyMass[i, 7])<> '') and (VarToStr(MyMass[i, 8])<> '') and (VarToStr(MyMass[i, 9])<> '') and
            (VarToStr(MyMass[i, 10])<> '') and (VarToStr(MyMass[i, 11])<> '') and (VarToStr(MyMass[i, 12])<> '')
          ) then
            begin
              SetLength(massMszFact,Length(massMszFact) + 1); //Добавляем ещё одну строку
              SetLength(massMszFact[High(massMszFact)], 13); //В добавленной строке (Последней по счёту) задаём размерность 13 столбцов
              massMszFact[High(massMszFact), massMSZFactSNILS] := MyMass[i, 2]; //Записываем название
              massMszFact[High(massMszFact), massMSZFactFamily] := MyMass[i, 3];   //Записываем код LMSZID
              massMszFact[High(massMszFact), massMSZFactName] := MyMass[i, 4];    //Записываем
              massMszFact[High(massMszFact), massMSZFactSurname] := MyMass[i, 5];    //Записываем
              massMszFact[High(massMszFact), massMSZFactGender] := MyMass[i, 6];    //Записываем
              massMszFact[High(massMszFact), massMSZFactBirthDate] := MyMass[i, 7];    //Записываем
              massMszFact[High(massMszFact), massMSZFactDecisionDate] := MyMass[i, 8];    //Записываем
              massMszFact[High(massMszFact), massMSZFactDateStart] := MyMass[i, 9];    //Записываем
              massMszFact[High(massMszFact), massMSZFactDateFinish] := MyMass[i, 10];    //Записываем
              massMszFact[High(massMszFact), massMSZFactLMSZName] := MyMass[i, 11];    //Записываем
              massMszFact[High(massMszFact), massMSZFactAmount] := MyMass[i, 12];    //Записываем
            end
          else
            Begin
              BadLinesFact := BadLinesFact + IntToStr(i) + ', ';
            End;
        end;

      //Выводим в логи количество найденных фактов
      RE_Log.Lines.Add(textIdent + 'Всего найдено фактов назначения МСЗ: ' + IntToStr(Length(massMszFact)));

      //Заполняем список ошибок
      if (BadLinesFact <> '') then
          M_Info.Lines.Add('Реестр фактов назначения МСЗ содержит некорректные строки (' + BadLinesFact + '); ');

  RE_Log.Lines.Add('');



  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + 'Чтение титульного листа');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // открытие листа книги
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Титульный'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['Титульный'];

      // выделение последней задействованной ячейки на листе
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // получение значений размера выбранного диапазона
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // присвоение массиву диапазона ячеек на листе
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //Задаём размер массива
      SetLength(orgInfo,8);
      //Вписываем данные
      orgInfo[orgInfoName] := MyMass[1, 2];
      orgInfo[orgInfoYear] := MyMass[2, 2];
      orgInfo[orgInfoMonth] := MyMass[3, 2];
      orgInfo[orgInfoFio] := MyMass[4, 2];
      orgInfo[orgInfoPhone] := MyMass[5, 2];
      orgInfo[orgInfoSpecialMarks] := MyMass[6, 2];
      orgInfo[orgInfoTemplateVersion] := MyMass[7, 2];

      //Выводим содержимое массива в логи
      RE_Log.Lines.Add(textIdent + 'Название организации: ' + orgInfo[orgInfoName]);
      RE_Log.Lines.Add(textIdent + 'Отчётный год: ' + orgInfo[orgInfoYear]);
      RE_Log.Lines.Add(textIdent + 'Отчётный месяц: ' + orgInfo[orgInfoMonth]);
      RE_Log.Lines.Add(textIdent + 'ФИО ответственного: ' + orgInfo[orgInfoFio]);
      RE_Log.Lines.Add(textIdent + 'Телефон ответственного: ' + orgInfo[orgInfoPhone]);
      RE_Log.Lines.Add(textIdent + 'Особые отметки: ' + orgInfo[orgInfoSpecialMarks]);
      RE_Log.Lines.Add(textIdent + 'Версия шаблона: ' + orgInfo[orgInfoTemplateVersion]);
      //RE_Log.Lines.Add(textIdent + 'Код ONMSZ орагнизации: ' + orgInfo[orgInfoONMSZCode]);

  RE_Log.Lines.Add('');


  //закрытие книги
  RE_Log.Lines.Add(textIdent + 'Закрытие файла.');
  ExcelApp.Quit;

  //отчистка переменных
  ExcelApp := Unassigned;
  ExcelSheet := Unassigned;
  MyMass := Unassigned;

  RE_Log.Lines.Add('');
  RE_Log.Lines.Add('');

End;


//Процедура проверяет данные об организации и подбирает код текущей организации
procedure TF_main.processingOrgInfo;
  var i : integer;
begin
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('Обработка титульного листа ');
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');

  //Проверка версии шаблона
  if (orgInfo[orgInfoTemplateVersion] <> VersionTemplate) then
    M_Error.Lines.Add('Версия открываемого файла-шаблона (' + orgInfo[orgInfoTemplateVersion] + ') и версия файла-шаблона для которого предназначена эта программа (' + VersionTemplate + ') не совпадают. ');

  //Подбор кода ONMSZ учреждения
  if (orgInfo[orgInfoName] <> '') then
    begin
      //Определяем код организации
      for i := Low(massOrgList) to High(massOrgList) do
        Begin
          if (massOrgList[i,massOrgListKratName] = orgInfo[orgInfoName]) then
            begin
              orgInfo[orgInfoONMSZCode] := massOrgList[i,massOrgListONMSZCode];
              break;
            end;
        End;
    end
  else
    begin
      M_Error.Lines.Add('На титульном листе не выбрана организация. ');
    end;

  if (orgInfo[orgInfoONMSZCode] = '') then
      M_Error.Lines.Add('Не удалось определить код ONMSZ выбранной организации. ')
  else
    RE_Log.Lines.Add(textIdent + 'Выбранной организации ' + orgInfo[orgInfoName] + ' назначен код ONMSZ ' + orgInfo[orgInfoONMSZCode] + ';');

  RE_Log.Lines.Add('');
  RE_Log.Lines.Add('');
end;


//Процедура проверяет реестр фактов назначения и подготавливает его для формирования в csv
procedure TF_main.processingMassMszFact;
  var i, j:integer;
begin
  //Запускаем перебор массива
  for i := Low(massMszFact) to High(massMszFact) do
    begin

      //Проверка СНИЛС
        //Удаляем все символы которые не цифры
        massMszFact[i, massMSZFactSNILS] := DeleteEverExNumber(massMszFact[i, massMSZFactSNILS]);
        //Проверяем что получилось в итоге в снилс
        if Length(massMszFact[i, massMSZFactSNILS]) <> 11 then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректный СНИЛС (' + massMszFact[i, massMSZFactSNILS] + ').');

      //Проверка фамилии
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactFamily] := Trim(massMszFact[i, massMSZFactFamily]);
        //Проверяем то что получилось в итоге от фамилии
        if (Length(massMszFact[i, massMSZFactFamily]) < 2) then  //Если осталось меньше двух символов
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректную фамилию (' + massMszFact[i, massMSZFactFamily] + ').');

      //Проверка имени
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactName] := Trim(massMszFact[i, massMSZFactName]);
        //Проверяем то что осталось в итоге от имени
        if (Length(massMszFact[i, massMSZFactName]) < 2) then  //Если осталось меньше двух символов
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректное имя (' + massMszFact[i, massMSZFactName] + ').');

      //Проверка отчества
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactSurname] := Trim(massMszFact[i, massMSZFactSurname]);
        //Проверяем то что осталось в итоге от имени
        if (Length(massMszFact[i, massMSZFactSurname]) < 2) then  //Если осталось меньше двух символов
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректное отчество (' + massMszFact[i, massMSZFactSurname] + ').');

      //Проверка пола
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactGender] := Trim(massMszFact[i, massMSZFactGender]);
        //Заменяем все "ж" на "Ж" и "м" на "М"
        massMszFact[i, massMSZFactGender] := ReplaceSexMark(massMszFact[i, massMSZFactGender] );
        if (massMszFact[i, massMSZFactGender] <> 'Ж') and (massMszFact[i, massMSZFactGender] <> 'М') then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректный пол (' + massMszFact[i, massMSZFactGender] + ').');

      //Проверка даты рождения
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactBirthDate] := Trim(massMszFact[i, massMSZFactBirthDate]);
        //Проверяем является ли введённое значение валидной датой
        if ( IsDate(massMszFact[i, massMSZFactBirthDate]) = false)  then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректную дату рождения (' + massMszFact[i, massMSZFactBirthDate] + ').');

      //Проверка даты назначения мсз
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactDecisionDate] := Trim(massMszFact[i, massMSZFactDecisionDate]);
        //Проверяем является ли введённое значение валидной датой
        if ( IsDate(massMszFact[i, massMSZFactDecisionDate]) = false)  then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректную дату принятия решения о назначении меры (' + massMszFact[i, massMSZFactDecisionDate] + ').');

      //Проверка даты начала действия мсз
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactDateStart] := Trim(massMszFact[i, massMSZFactDateStart]);
        //Проверяем является ли введённое значение валидной датой
        if ( IsDate(massMszFact[i, massMSZFactDateStart]) = false)  then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректную дату начала действия меры (' + massMszFact[i, massMSZFactDateStart] + ').');

      //Проверка даты конца действия мсз
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactDateFinish] := Trim(massMszFact[i, massMSZFactDateFinish]);
        //Проверяем является ли введённое значение валидной датой
        if ( IsDate(massMszFact[i, massMSZFactDateFinish]) = false)  then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректную дату окончания действия меры (' + massMszFact[i, massMSZFactDateFinish] + ').');


      //Проверка названия меры
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactLMSZName] := Trim(massMszFact[i, massMSZFactLMSZName]);
        //Проверяем то что осталось
        if (massMszFact[i, massMSZFactLMSZName] = '') then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректное название меры (' + massMszFact[i, massMSZFactLMSZName] + ').');

      //Подбор кодов меры
        //Перебираем справочник мер и подбираем коды
        for j := Low(massMszList) to High(massMszList) do
          begin
            if massMszFact[i, massMSZFactLMSZName] = massMszList[j,massMszListKratName] then
              Begin
                massMszFact[i, massMSZFactLMSZID] := massMszList[j,massMszListLMSZID];
                massMszFact[i, massMSZFactLMSZCategoryID] := massMszList[j,massMszListСategoryID];
                Break;
              End;
          end;
      //Проверяем нашлось ли что в итоге
        if (massMszFact[i, massMSZFactLMSZID] = '') then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' не удалось определить идентификатор ЛМСЗ меры (' + massMszFact[i, massMSZFactLMSZID] + ').');
        if (massMszFact[i, massMSZFactLMSZCategoryID] = '') then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' не удалось определить идентификатор локальной категории меры (' + massMszFact[i, massMSZFactLMSZCategoryID] + ').');




      //Проверка суммы меры
        //Удаляем пробелы в начале и в конце
        massMszFact[i, massMSZFactAmount] := Trim(massMszFact[i, massMSZFactAmount]);
        //Заменяем все "." на ","
        massMszFact[i, massMSZFactAmount] := ReplaceNumberSeparator(massMszFact[i, massMSZFactAmount]);
        //Проверяем является ли то что осталось числом float
        if (IsFloat(massMszFact[i, massMSZFactAmount]) = false) then
          M_Error.Lines.Add('Строка ' + IntToStr(i + 1) + ' содержит некорректную сумму (' + massMszFact[i, massMSZFactAmount] + ').')
        else //Если сумма корректна - накапливаем итого по каждой мере в справочнике мер
          Begin
            //Перебираем справочник мер и подбираем коды
            for j := Low(massMszList) to High(massMszList) do
              begin
                if massMszFact[i, massMSZFactLMSZName] = massMszList[j,massMszListKratName] then
                  Begin
                    //Накапливаем сумму по каждой мере
                    massMszList[j,massMszListSumm] := FloatToStr(StrToFloat(massMszList[j,massMszListSumm]) + StrToFloat(massMszFact[i, massMSZFactAmount]));
                    Break;
                  End;
              end;
          End;

    end;

  //Выводим полученные данные в логи
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('Обработка реестра фактов назначения МСЗ');
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add(textIdent + 'Всего суммы по мерам:');
  for i := Low(massMszList) to High(massMszList) do
    begin
      RE_Log.Lines.Add(textIdent + IntToStr(i) + ') ' + massMszList[i, massMszListKratName] + ': ' + massMszList[i, massMszListSumm] + 'руб.');
    end;

  RE_Log.Lines.Add('');
  RE_Log.Lines.Add('');
end;



procedure TF_main.B_GridClick(Sender: TObject);
begin
  P_Info.Visible := False;
  RE_Log.Visible := False;
  SG_Fact.Visible := True;
end;

procedure TF_main.B_InfoClick(Sender: TObject);
begin
  P_Info.Visible := True;
  RE_Log.Visible := False;
  SG_Fact.Visible := False;
end;

procedure TF_main.B_logClick(Sender: TObject);
begin
  P_Info.Visible := false;
  RE_Log.Visible := true;
  SG_Fact.Visible := False;
end;

procedure TF_main.B_OpenFileClick(Sender: TObject);
begin
  if OD_OpenXlsxFile.Execute then
  Begin
    //Очищаем всё перед работой
    M_Info.Clear;
    M_Error.Clear;
    RE_Log.Clear;

    //Обнуляем таблицы
    SetLength(massMszList, 0);
    SetLength(massMszFact, 0);
    SetLength(massOrgList, 0);
    SetLength(orgInfo, 0);

    //Пишем дату и начальную инфу
    RE_Log.Lines.Add('');
    RE_Log.Lines.Add(DateTimeToStr(now));
    RE_Log.Lines.Add('');

    //Определяем данные файла
    XlsxFile := OD_OpenXlsxFile.FileName;
    XlsxFileDir := ExtractFileDir(XlsxFile);
    XlsxFileName := ExtractFileName(XlsxFile);
    XlsxFileNameNoExt := StringReplace(ExtractFileName(XlsxFile),ExtractFileExt(XlsxFile),'',[]);
        //Отладочный вывод путей
        {RE_Log.Lines.Add('Путь к файлу: ' + XlsxFile);
        RE_Log.Lines.Add('Путь папке файла: ' + XlsxFileDir);
        RE_Log.Lines.Add('Имя файла: ' + XlsxFileName);
        RE_Log.Lines.Add('Имя файла без расширения: ' + XlsxFileNameNoExt); }

    //Запускаем получение данных
    getData(XlsxFile);

    //Запускаем обработку титульного листа
    if (M_Error.Text = '') then   //Если в предыдущих шагах ошибок не найдено...
        processingOrgInfo();

    //Запускаем обработку реестра фактов
    if (M_Error.Text = '') then   //Если в предыдущих шагах ошибок не найдено...
      processingMassMszFact();


    //Ренедерим данные в табличку
    renderTableMszFact();

    //Запускаем сохранение в файл CSV
    if (M_Error.Text = '') then   //Если в предыдущих шагах ошибок не найдено...
      renderDataToCsv();


    //Вывод сообщений
    if (M_Info.Text <> '') then
      Begin
        RE_Log.Lines.Add('');
        RE_Log.Lines.Add('Обнаруженные НЕ критические ошибки:');
        RE_Log.Text := RE_Log.Text + M_Info.Text;
      End;

    //Вывод ошибок
    if (M_Error.Text <> '') then
      Begin
        RE_Log.Lines.Add('');
        RE_Log.Lines.Add('Обнаруженные КРИТИЧЕСКИЕ ошибки:');
        RE_Log.Text := RE_Log.Text + M_Error.Text;
      End;

    B_log.Click; //Открываем форму с логами если она не была открыта

  End;

end;

procedure TF_main.FormCreate(Sender: TObject);
begin
  //Применяем первоначальные настройки
    //Оформление заголовка окна
    F_main.Caption := 'Конвертер фактов назначения МСЗ xlsx->csv.  Версия программы ' + VersionSoftware + '  Версия шаблона ' + VersionTemplate;

    //Оформление раздела "О программе"
    L_VersionSoftware.caption := VersionSoftware;
    L_VersionTemplate.Caption := VersionTemplate;
    L_DateBuild.Caption := DateBuild;

    //Задание отступов у RE_log
    RE_Log.Paragraph.FirstIndent:=10;
    RE_Log.Paragraph.LeftIndent:=10;

    //Включаем интерфейс на страницу логов
    B_log.Click;

end;

end.
