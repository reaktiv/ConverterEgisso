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

    procedure renderTableMszFact(); //��������� ���������� ������ ������ ���������� ���
    procedure renderDataToCsv(); //����������� ��� ������ �� massMszFact � ������ CSV


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

  //�������
  massMszList : array of array of String; //������ �������� ������ ��������� ��� ���. ���������
  massMszFact : array of array of String; //������ �������� ������ ������ ���������� ��� ���. ���������
  massOrgList : array of array of String; //������ �������� ������ �����������
  orgInfo : array of String; //������ �������� ������ ����������

  XlsxFile : String; //������ ���� (� ������) � ������������ �����-�������
  XlsxFileDir : String; //���� � ����� � ����������� ������-��������
  XlsxFileName : String; //�������� ����� � ����������� (����� �����)
  XlsxFileNameNoExt : String; //�������� ����� ��� ����������

  M_Info : TMemo; //������ ������ �� ����������� ������
  M_Error : TMemo; //������ ������ ����������� ������

const

  //�������� ����������
  VersionSoftware = '1.0'; //������ ���������
  VersionTemplate = '2.0'; //������ ������� � ������� ����� �������� ���������
  DateBuild = '05.12.2020'; //���� ������

  //���������� ������� massOrgList
  massOrgListKratName = 0;
  massOrgListONMSZCode = 1;

  //���������� ������� massMszList
  massMszListKratName = 0;
  massMszListLMSZID = 1;
  massMszList�ategoryID = 2;
  massMszListSumm = 3;

  //���������� ������� massMSZFact
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
  massMSZFactLMSZID = 11; //����� ������� LMSZID ����
  massMSZFactLMSZCategoryID= 12; //����� ������� �ategoryID ����

  //���������� ������� orgInfo
  orgInfoName = 0;
  orgInfoYear = 1;
  orgInfoMonth = 2;
  orgInfoFio = 3;
  orgInfoPhone = 4;
  orgInfoSpecialMarks = 5;
  orgInfoTemplateVersion = 6;
  orgInfoONMSZCode = 7; //����� ������� ��� ��������� �����������

  //������ � �����
  textIdent = '     '; //������ ������ � ��������

implementation

{$R *.dfm}

//������� ������� �� ���������� ������ �� ����� ����
function TF_main.DeleteEverExNumber(Text:String) : String;
  var i :integer;
  TextOut : String;
Begin
  TextOut := '';
  for i := 1 to length(Text) do
    if Text[i] in ['0'..'9'] then TextOut := TextOut + Text[i];
  result := TextOut;
End;


//������� �������� ��� "�" �� "�" � ��� "�" �� "�"
function TF_main.ReplaceSexMark(Text:String): String;
Begin
  Text := StringReplace(Text, '�', '�', [rfReplaceAll]);
  Text := StringReplace(Text, '�', '�', [rfReplaceAll]);
  Result := Text;
End;


//������� �������� ��� "." �� ","
function TF_main.ReplaceNumberSeparator(Text:String) : String;
Begin
  Text := StringReplace(Text, '.', ',', [rfReplaceAll]);
  Result := Text;
End;


//������� ��������� �������� �� ���������� ������ �����
function TF_main.IsDate(Text:String): Boolean;
  var dat:TDateTime;
Begin
  if TryStrToDate(Text,dat) then Result := true else Result := false;
End;


//������� ��������� �������� �� ���������� ������ ����� float
function TF_main.IsFloat(Text:String):Boolean;
  Var Mu:double;
begin
  if TryStrToFloat(Text, Mu) then Result := true else Result := false;
End;


//��������� ���������� � ������� ������ ������ ���������� ���
procedure TF_Main.renderTableMszFact();
  var i,j : integer;
Begin
  //������������ ������
  if (Length(massMszFact) > 0)  then
    begin

      //�������������� ����
      SG_Fact.ColCount :=  14;
      SG_Fact.RowCount :=  Length(massMszFact) + 1;

      SG_Fact.FixedCols := 1;
      SG_Fact.FixedRows := 1;

      //��������� ������ � ��������� ��������
        //������� �������
        SG_Fact.ColWidths[0] := 60;
        SG_Fact.Cells[0,0] := '� �.�';
        //������� �����
        SG_Fact.ColWidths[1] := 130;
        SG_Fact.Cells[1,0] := '�����';
        //������� �������
        SG_Fact.ColWidths[2] := 150;
        SG_Fact.Cells[2,0] := '�������';
        //������� ���
        SG_Fact.ColWidths[3] := 150;
        SG_Fact.Cells[3,0] := '���';
        //������� ��������
        SG_Fact.ColWidths[4] := 150;
        SG_Fact.Cells[4,0] := '��������';
        //������� ���
        SG_Fact.ColWidths[5] := 50;
        SG_Fact.Cells[5,0] := '���';
        //������� ���� ��������
        SG_Fact.ColWidths[6] := 120;
        SG_Fact.Cells[6,0] := '���� ��������';
        //������� ���� ������� � ����������
        SG_Fact.ColWidths[7] := 120;
        SG_Fact.Cells[7,0] := '���� �������';
        //������� ���� ������ ��������
        SG_Fact.ColWidths[8] := 120;
        SG_Fact.Cells[8,0] := '���� ������';
        //������� ���� ��������� ��������
        SG_Fact.ColWidths[9] := 120;
        SG_Fact.Cells[9,0] := '���� ���������';
        //������� ����
        SG_Fact.ColWidths[10] := 200;
        SG_Fact.Cells[10,0] := '����';
        //������� �����
        SG_Fact.ColWidths[11] := 100;
        SG_Fact.Cells[11,0] := '�����';
        //������� ������������� ����
        SG_Fact.ColWidths[12] := 300;
        SG_Fact.Cells[12,0] := '������������� ����';
        //������� ������������� ����
        SG_Fact.ColWidths[13] := 300;
        SG_Fact.Cells[13,0] := '������������� ��������� ���������';


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


//��������� ����������� ������ �� massMSZFact � ������ CSV
procedure TF_Main.renderDataToCsv();
  var i:Integer;
      delimiter : String;
begin
  //�������������� Memo
  M_CSV.Clear;
  delimiter := ';';

  //����� ��������� � ���
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('������� ������');
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');

  //���������� ������ ������
  M_CSV.Lines.Add('RecType;LMSZID;categoryID;ONMSZCode;SNILS_recip;FamilyName_recip;Name_recip;' +
  'Patronymic_recip;Gender_recip;BirthDate_recip;doctype_recip;doc_Series_recip;doc_Number_recip;' +
  'doc_IssueDate_recip;doc_Issuer_recip;SNILS_reason;FamilyName_reason;Name_reason;Patronymic_reason;' +
  'Gender_reason;BirthDate_reason;doctype_reason;doc_Series_reason;doc_Number_reason;doc_IssueDate_reason;' +
  'doc_Issuer_reason;decision_date;dateStart;dateFinish;usingSign;criteria;FormCode;amount;measuryCode;' +
  'monetization;content;comment;equivalentAmount');

  //����������� ������ � �������
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
      delimiter + delimiter + delimiter + delimiter + delimiter +   //���������� 17 ����� � �������
      delimiter + delimiter + delimiter + delimiter + delimiter +
      delimiter + delimiter + delimiter + delimiter + delimiter +
      delimiter + delimiter +
      massMszFact[i, massMSZFactDecisionDate] + delimiter +
      massMszFact[i, massMSZFactDateStart] + delimiter +
      massMszFact[i, massMSZFactDateFinish] + delimiter +
      '���' + delimiter + delimiter + '01' + delimiter +
      massMszFact[i, massMSZFactAmount] + delimiter +
      '1' + delimiter + '0' + delimiter + delimiter + delimiter);
    end;

  //��������� ������ � csv ����
  if (DirectoryExists(XlsxFileDir + '\csv') <> true) then
    ForceDirectories(XlsxFileDir + '\csv'); //���� ������ ����� ��� - ������.

  if (DirectoryExists(XlsxFileDir + '\csv') <> true) then
    M_Error.Lines.Add('�� ������ ������� ����� ' + XlsxFileDir + '\csv ;'); //����� ��� ��� �� ���������� ������� �����

  if (M_Error.Text = '') then //���� �� ���������� ������ �� ���� ����������� ������
    Begin
      //������� ���� ���� �� ����������
      if (FileExists(XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv') = true) then
        DeleteFile((XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv'));

      if (FileExists(XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv') <> true) then
        Begin
          M_CSV.Lines.SaveToFile(XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv');
          RE_Log.Lines.Add(textIdent + '������ ��������� � ���� ' + XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv; ');
        end
      else
        M_Error.Lines.Add('�� ������ ������� ���� ' + XlsxFileDir + '\csv\' + XlsxFileNameNoExt + '.csv'); //����� ��� ��� �� ���������� ������� ����
    End;



  //��������� ���� � �����-����
  if (DirectoryExists(XlsxFileDir + '\log') <> true) then
    ForceDirectories(XlsxFileDir + '\log'); //���� ������ ����� ��� - ������.

  if (DirectoryExists(XlsxFileDir + '\log') <> true) then
    M_Error.Lines.Add('�� ������ ������� ����� ' + XlsxFileDir + '\log ;'); //����� ��� ��� �� ���������� ������� �����

  if (M_Error.Text = '') then //���� �� ���������� ������ �� ���� ����������� ������
    Begin
      //������� ���� ���� �� ����������
      if (FileExists(XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log') = true) then
        DeleteFile((XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log'));

      if (FileExists(XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log') <> true) then
        Begin
          RE_Log.Lines.Add(textIdent + '���� ��������� � ���� ' + XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log; ');
          RE_Log.Lines.SaveToFile(XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log');
        end
      else
        M_Error.Lines.Add('�� ������ ������� ���� ' + XlsxFileDir + '\log\' + XlsxFileNameNoExt + '.log'); //����� ��� ��� �� ���������� ������� ����
    End;

end;


//��������� ��������� ������ � ������� �� ���������� �����
procedure TF_main.getData(xlsxFilePath:String);
  const
    xlCellTypeLastCell = $0000000B;
  var
    //��������� ��� ������ � ������ xlsx
    ExcelApp, ExcelSheet: OLEVariant;
    MyMass: Variant;
    i, x, y: Integer;

    BadLinesMsz : String; //������ ������ ���������� ����������� ����� ����������� ���
    BadLinesOrg : String; //������ ������ ���������� ����������� ����� ����������� ����������
    BadLinesFact : String; //������ ������ ���������� ����������� ����� ������� ������ ���������� ���

Begin
  //���������� ����������
  BadLinesMsz := '';
  BadLinesOrg := '';
  BadLinesFact := '';


  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('�������� ������ �� ����� ' + xlsxFilePath);
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');



  //�������� OLE-������� Excel
  RE_Log.Lines.Add(textIdent + '�������� �����...');
  ExcelApp := CreateOleObject('Excel.Application');
  //�������� ����� Excel
  RE_Log.Lines.Add(textIdent + '�������� �����...');
  ExcelApp.Workbooks.Open(xlsxFilePath);


  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + '������ ����������� �����������');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // �������� ����� �����
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['���.�����������'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['���.�����������'];

      // ��������� ��������� ��������������� ������ �� �����
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // ��������� �������� ������� ���������� ���������
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // ���������� ������� ��������� ����� �� �����
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //��������� ������ � ����� ������
      SetLength(massOrgList, 0); //����� ������� ����������� �������
      for i := 2 to x do  //������� �� ������ ������ ���������� �� ��� ��������
        begin
          if ((VarToStr(MyMass[i, 1])<> '') and (VarToStr(MyMass[i, 2])<> '') and (VarToStr(MyMass[i, 3])<> '')) then  //���� ������ � ������ � ������ ������� ���������
            Begin
              SetLength(massOrgList,Length(massOrgList) + 1); //��������� ��� ���� ������
              SetLength(massOrgList[High(massOrgList)], 2); //� ����������� ������ (��������� �� �����) ����� ����������� 2 �������
              massOrgList[High(massOrgList), massOrgListKratName] := MyMass[i, 2]; //���������� ��������
              massOrgList[High(massOrgList), massOrgListONMSZCode] := MyMass[i, 3];   //���������� ��� LMSZID
            End
          else
            Begin
              BadLinesOrg := BadLinesOrg + IntToStr(i) + ', ';
            End;

        end;

      //������� � ���� ���������� ��������� ����������
      RE_Log.Lines.Add(textIdent + '����� ������� ����������: ' + IntToStr(Length(massOrgList)));

      //������� ������ ����������
      for i := Low(massOrgList) to High(massOrgList) do
        Begin
          RE_Log.Lines.Add(textIdent + IntToStr(i+1) + ')   ' + massOrgList[i,massOrgListKratName] + '   (ONMSZCode: ' +  massOrgList[i,massOrgListONMSZCode] + ');');
        end;

      //��������� ������ ������
      if (BadLinesOrg <> '') then
          M_Info.Lines.Add('���������� ����������� �������� ������������ ������ (' + BadLinesOrg + '); ');

  RE_Log.Lines.Add('');



  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + '������ ����������� ���');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // �������� ����� �����
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['���.����'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['���.����'];

      // ��������� ��������� ��������������� ������ �� �����
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // ��������� �������� ������� ���������� ���������
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // ���������� ������� ��������� ����� �� �����
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //�������������� ��������� ������ � ���������� ������
      SetLength(massMszList,0); //����� ������� ����������� �������
      for i := 2 to x do  //������� �� ������ ������ ���������� �� ��� ��������
        begin
          if ((VarToStr(MyMass[i, 1])<> '') and (VarToStr(MyMass[i, 2])<> '') and (VarToStr(MyMass[i, 3])<> '') and (VarToStr(MyMass[i, 4])<> '')) then
            begin
              SetLength(massMszList,Length(massMszList) + 1); //��������� ��� ���� ������
              SetLength(massMszList[High(massMszList)], 4); //� ����������� ������ (��������� �� �����) ����� ����������� 4 �������
              massMszList[High(massMszList), massMszListKratName] := MyMass[i, 2]; //���������� ��������
              massMszList[High(massMszList), massMszListLmszID] := MyMass[i, 3];   //���������� ��� LMSZID
              massMszList[High(massMszList), massMszList�ategoryID] := MyMass[i, 4];    //���������� ��� categoryID
              massMszList[High(massMszList), massMszListSumm] := '0'; //���������� ���� ����� ����� ��������� ������������� ��� � ����� � ����������� �����
            end
          else
            Begin
              BadLinesMsz := BadLinesMsz + IntToStr(i) + ', ';
            End;
        end;

      //������� � ���� ���������� ��������� ���
      RE_Log.Lines.Add(textIdent + '����� ������� ���: ' + IntToStr(Length(massMszList)));

      //������� � ���� ������ ��������� ���
      for i := Low(massMszList) to High(massMszList) do
        Begin
          RE_Log.Lines.Add(textIdent + IntToStr(i+1) + ')   '+  massMszList[i,massMszListKratName] + '   (LMSZID:' +  massMszList[i,massMszListLmszID] + ', categoryID:' + massMszList[i,massMszList�ategoryID] + ');');
        end;

      //��������� ������ ������
      if (BadLinesMsz <> '') then
          M_Info.Lines.Add('���������� ��� �������� ������������ ������ (' + BadLinesMsz + '); ');

  RE_Log.Lines.Add('');



  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + '������ ������� ������ ���������� ���');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // �������� ����� �����
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['������ ������ ���������� ���'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['������ ������ ���������� ���'];

      // ��������� ��������� ��������������� ������ �� �����
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // ��������� �������� ������� ���������� ���������
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // ���������� ������� ��������� ����� �� �����
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //�������������� ��������� ������ � ���������� ������
      SetLength(massMszFact,0); //����� ������� ����������� �������
      for i := 2 to x do  //������� �� ������ ������ ���������� �� ��� ��������
        begin
          if (
            (VarToStr(MyMass[i, 1])<> '') and (VarToStr(MyMass[i, 2])<> '') and (VarToStr(MyMass[i, 3])<> '') and
            (VarToStr(MyMass[i, 4])<> '') and (VarToStr(MyMass[i, 5])<> '') and (VarToStr(MyMass[i, 6])<> '') and
            (VarToStr(MyMass[i, 7])<> '') and (VarToStr(MyMass[i, 8])<> '') and (VarToStr(MyMass[i, 9])<> '') and
            (VarToStr(MyMass[i, 10])<> '') and (VarToStr(MyMass[i, 11])<> '') and (VarToStr(MyMass[i, 12])<> '')
          ) then
            begin
              SetLength(massMszFact,Length(massMszFact) + 1); //��������� ��� ���� ������
              SetLength(massMszFact[High(massMszFact)], 13); //� ����������� ������ (��������� �� �����) ����� ����������� 13 ��������
              massMszFact[High(massMszFact), massMSZFactSNILS] := MyMass[i, 2]; //���������� ��������
              massMszFact[High(massMszFact), massMSZFactFamily] := MyMass[i, 3];   //���������� ��� LMSZID
              massMszFact[High(massMszFact), massMSZFactName] := MyMass[i, 4];    //����������
              massMszFact[High(massMszFact), massMSZFactSurname] := MyMass[i, 5];    //����������
              massMszFact[High(massMszFact), massMSZFactGender] := MyMass[i, 6];    //����������
              massMszFact[High(massMszFact), massMSZFactBirthDate] := MyMass[i, 7];    //����������
              massMszFact[High(massMszFact), massMSZFactDecisionDate] := MyMass[i, 8];    //����������
              massMszFact[High(massMszFact), massMSZFactDateStart] := MyMass[i, 9];    //����������
              massMszFact[High(massMszFact), massMSZFactDateFinish] := MyMass[i, 10];    //����������
              massMszFact[High(massMszFact), massMSZFactLMSZName] := MyMass[i, 11];    //����������
              massMszFact[High(massMszFact), massMSZFactAmount] := MyMass[i, 12];    //����������
            end
          else
            Begin
              BadLinesFact := BadLinesFact + IntToStr(i) + ', ';
            End;
        end;

      //������� � ���� ���������� ��������� ������
      RE_Log.Lines.Add(textIdent + '����� ������� ������ ���������� ���: ' + IntToStr(Length(massMszFact)));

      //��������� ������ ������
      if (BadLinesFact <> '') then
          M_Info.Lines.Add('������ ������ ���������� ��� �������� ������������ ������ (' + BadLinesFact + '); ');

  RE_Log.Lines.Add('');



  RE_Log.Lines.Add('');
  RE_Log.Lines.Add(textIdent + '������ ���������� �����');
  RE_Log.Lines.Add(textIdent + '------------------------------------------------------------------------------');

      // �������� ����� �����
      ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['���������'].Activate;
      ExcelSheet := ExcelApp.Workbooks[ExtractFileName(xlsxFilePath)].WorkSheets['���������'];

      // ��������� ��������� ��������������� ������ �� �����
      ExcelSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;

      // ��������� �������� ������� ���������� ���������
      x := ExcelApp.ActiveCell.Row;
      y := ExcelApp.ActiveCell.Column;

      // ���������� ������� ��������� ����� �� �����
      MyMass := ExcelApp.Range['A1', ExcelApp.Cells.Item[X, Y]].Value;

      //����� ������ �������
      SetLength(orgInfo,8);
      //��������� ������
      orgInfo[orgInfoName] := MyMass[1, 2];
      orgInfo[orgInfoYear] := MyMass[2, 2];
      orgInfo[orgInfoMonth] := MyMass[3, 2];
      orgInfo[orgInfoFio] := MyMass[4, 2];
      orgInfo[orgInfoPhone] := MyMass[5, 2];
      orgInfo[orgInfoSpecialMarks] := MyMass[6, 2];
      orgInfo[orgInfoTemplateVersion] := MyMass[7, 2];

      //������� ���������� ������� � ����
      RE_Log.Lines.Add(textIdent + '�������� �����������: ' + orgInfo[orgInfoName]);
      RE_Log.Lines.Add(textIdent + '�������� ���: ' + orgInfo[orgInfoYear]);
      RE_Log.Lines.Add(textIdent + '�������� �����: ' + orgInfo[orgInfoMonth]);
      RE_Log.Lines.Add(textIdent + '��� ��������������: ' + orgInfo[orgInfoFio]);
      RE_Log.Lines.Add(textIdent + '������� ��������������: ' + orgInfo[orgInfoPhone]);
      RE_Log.Lines.Add(textIdent + '������ �������: ' + orgInfo[orgInfoSpecialMarks]);
      RE_Log.Lines.Add(textIdent + '������ �������: ' + orgInfo[orgInfoTemplateVersion]);
      //RE_Log.Lines.Add(textIdent + '��� ONMSZ �����������: ' + orgInfo[orgInfoONMSZCode]);

  RE_Log.Lines.Add('');


  //�������� �����
  RE_Log.Lines.Add(textIdent + '�������� �����.');
  ExcelApp.Quit;

  //�������� ����������
  ExcelApp := Unassigned;
  ExcelSheet := Unassigned;
  MyMass := Unassigned;

  RE_Log.Lines.Add('');
  RE_Log.Lines.Add('');

End;


//��������� ��������� ������ �� ����������� � ��������� ��� ������� �����������
procedure TF_main.processingOrgInfo;
  var i : integer;
begin
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('��������� ���������� ����� ');
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');

  //�������� ������ �������
  if (orgInfo[orgInfoTemplateVersion] <> VersionTemplate) then
    M_Error.Lines.Add('������ ������������ �����-������� (' + orgInfo[orgInfoTemplateVersion] + ') � ������ �����-������� ��� �������� ������������� ��� ��������� (' + VersionTemplate + ') �� ���������. ');

  //������ ���� ONMSZ ����������
  if (orgInfo[orgInfoName] <> '') then
    begin
      //���������� ��� �����������
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
      M_Error.Lines.Add('�� ��������� ����� �� ������� �����������. ');
    end;

  if (orgInfo[orgInfoONMSZCode] = '') then
      M_Error.Lines.Add('�� ������� ���������� ��� ONMSZ ��������� �����������. ')
  else
    RE_Log.Lines.Add(textIdent + '��������� ����������� ' + orgInfo[orgInfoName] + ' �������� ��� ONMSZ ' + orgInfo[orgInfoONMSZCode] + ';');

  RE_Log.Lines.Add('');
  RE_Log.Lines.Add('');
end;


//��������� ��������� ������ ������ ���������� � �������������� ��� ��� ������������ � csv
procedure TF_main.processingMassMszFact;
  var i, j:integer;
begin
  //��������� ������� �������
  for i := Low(massMszFact) to High(massMszFact) do
    begin

      //�������� �����
        //������� ��� ������� ������� �� �����
        massMszFact[i, massMSZFactSNILS] := DeleteEverExNumber(massMszFact[i, massMSZFactSNILS]);
        //��������� ��� ���������� � ����� � �����
        if Length(massMszFact[i, massMSZFactSNILS]) <> 11 then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ����� (' + massMszFact[i, massMSZFactSNILS] + ').');

      //�������� �������
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactFamily] := Trim(massMszFact[i, massMSZFactFamily]);
        //��������� �� ��� ���������� � ����� �� �������
        if (Length(massMszFact[i, massMSZFactFamily]) < 2) then  //���� �������� ������ ���� ��������
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ������� (' + massMszFact[i, massMSZFactFamily] + ').');

      //�������� �����
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactName] := Trim(massMszFact[i, massMSZFactName]);
        //��������� �� ��� �������� � ����� �� �����
        if (Length(massMszFact[i, massMSZFactName]) < 2) then  //���� �������� ������ ���� ��������
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ��� (' + massMszFact[i, massMSZFactName] + ').');

      //�������� ��������
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactSurname] := Trim(massMszFact[i, massMSZFactSurname]);
        //��������� �� ��� �������� � ����� �� �����
        if (Length(massMszFact[i, massMSZFactSurname]) < 2) then  //���� �������� ������ ���� ��������
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ �������� (' + massMszFact[i, massMSZFactSurname] + ').');

      //�������� ����
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactGender] := Trim(massMszFact[i, massMSZFactGender]);
        //�������� ��� "�" �� "�" � "�" �� "�"
        massMszFact[i, massMSZFactGender] := ReplaceSexMark(massMszFact[i, massMSZFactGender] );
        if (massMszFact[i, massMSZFactGender] <> '�') and (massMszFact[i, massMSZFactGender] <> '�') then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ��� (' + massMszFact[i, massMSZFactGender] + ').');

      //�������� ���� ��������
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactBirthDate] := Trim(massMszFact[i, massMSZFactBirthDate]);
        //��������� �������� �� �������� �������� �������� �����
        if ( IsDate(massMszFact[i, massMSZFactBirthDate]) = false)  then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ���� �������� (' + massMszFact[i, massMSZFactBirthDate] + ').');

      //�������� ���� ���������� ���
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactDecisionDate] := Trim(massMszFact[i, massMSZFactDecisionDate]);
        //��������� �������� �� �������� �������� �������� �����
        if ( IsDate(massMszFact[i, massMSZFactDecisionDate]) = false)  then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ���� �������� ������� � ���������� ���� (' + massMszFact[i, massMSZFactDecisionDate] + ').');

      //�������� ���� ������ �������� ���
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactDateStart] := Trim(massMszFact[i, massMSZFactDateStart]);
        //��������� �������� �� �������� �������� �������� �����
        if ( IsDate(massMszFact[i, massMSZFactDateStart]) = false)  then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ���� ������ �������� ���� (' + massMszFact[i, massMSZFactDateStart] + ').');

      //�������� ���� ����� �������� ���
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactDateFinish] := Trim(massMszFact[i, massMSZFactDateFinish]);
        //��������� �������� �� �������� �������� �������� �����
        if ( IsDate(massMszFact[i, massMSZFactDateFinish]) = false)  then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ���� ��������� �������� ���� (' + massMszFact[i, massMSZFactDateFinish] + ').');


      //�������� �������� ����
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactLMSZName] := Trim(massMszFact[i, massMSZFactLMSZName]);
        //��������� �� ��� ��������
        if (massMszFact[i, massMSZFactLMSZName] = '') then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ �������� ���� (' + massMszFact[i, massMSZFactLMSZName] + ').');

      //������ ����� ����
        //���������� ���������� ��� � ��������� ����
        for j := Low(massMszList) to High(massMszList) do
          begin
            if massMszFact[i, massMSZFactLMSZName] = massMszList[j,massMszListKratName] then
              Begin
                massMszFact[i, massMSZFactLMSZID] := massMszList[j,massMszListLMSZID];
                massMszFact[i, massMSZFactLMSZCategoryID] := massMszList[j,massMszList�ategoryID];
                Break;
              End;
          end;
      //��������� ������� �� ��� � �����
        if (massMszFact[i, massMSZFactLMSZID] = '') then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �� ������� ���������� ������������� ���� ���� (' + massMszFact[i, massMSZFactLMSZID] + ').');
        if (massMszFact[i, massMSZFactLMSZCategoryID] = '') then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �� ������� ���������� ������������� ��������� ��������� ���� (' + massMszFact[i, massMSZFactLMSZCategoryID] + ').');




      //�������� ����� ����
        //������� ������� � ������ � � �����
        massMszFact[i, massMSZFactAmount] := Trim(massMszFact[i, massMSZFactAmount]);
        //�������� ��� "." �� ","
        massMszFact[i, massMSZFactAmount] := ReplaceNumberSeparator(massMszFact[i, massMSZFactAmount]);
        //��������� �������� �� �� ��� �������� ������ float
        if (IsFloat(massMszFact[i, massMSZFactAmount]) = false) then
          M_Error.Lines.Add('������ ' + IntToStr(i + 1) + ' �������� ������������ ����� (' + massMszFact[i, massMSZFactAmount] + ').')
        else //���� ����� ��������� - ����������� ����� �� ������ ���� � ����������� ���
          Begin
            //���������� ���������� ��� � ��������� ����
            for j := Low(massMszList) to High(massMszList) do
              begin
                if massMszFact[i, massMSZFactLMSZName] = massMszList[j,massMszListKratName] then
                  Begin
                    //����������� ����� �� ������ ����
                    massMszList[j,massMszListSumm] := FloatToStr(StrToFloat(massMszList[j,massMszListSumm]) + StrToFloat(massMszFact[i, massMSZFactAmount]));
                    Break;
                  End;
              end;
          End;

    end;

  //������� ���������� ������ � ����
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add('��������� ������� ������ ���������� ���');
  RE_Log.Lines.Add('------------------------------------------------------------------------------------------------------------------');
  RE_Log.Lines.Add(textIdent + '����� ����� �� �����:');
  for i := Low(massMszList) to High(massMszList) do
    begin
      RE_Log.Lines.Add(textIdent + IntToStr(i) + ') ' + massMszList[i, massMszListKratName] + ': ' + massMszList[i, massMszListSumm] + '���.');
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
    //������� �� ����� �������
    M_Info.Clear;
    M_Error.Clear;
    RE_Log.Clear;

    //�������� �������
    SetLength(massMszList, 0);
    SetLength(massMszFact, 0);
    SetLength(massOrgList, 0);
    SetLength(orgInfo, 0);

    //����� ���� � ��������� ����
    RE_Log.Lines.Add('');
    RE_Log.Lines.Add(DateTimeToStr(now));
    RE_Log.Lines.Add('');

    //���������� ������ �����
    XlsxFile := OD_OpenXlsxFile.FileName;
    XlsxFileDir := ExtractFileDir(XlsxFile);
    XlsxFileName := ExtractFileName(XlsxFile);
    XlsxFileNameNoExt := StringReplace(ExtractFileName(XlsxFile),ExtractFileExt(XlsxFile),'',[]);
        //���������� ����� �����
        {RE_Log.Lines.Add('���� � �����: ' + XlsxFile);
        RE_Log.Lines.Add('���� ����� �����: ' + XlsxFileDir);
        RE_Log.Lines.Add('��� �����: ' + XlsxFileName);
        RE_Log.Lines.Add('��� ����� ��� ����������: ' + XlsxFileNameNoExt); }

    //��������� ��������� ������
    getData(XlsxFile);

    //��������� ��������� ���������� �����
    if (M_Error.Text = '') then   //���� � ���������� ����� ������ �� �������...
        processingOrgInfo();

    //��������� ��������� ������� ������
    if (M_Error.Text = '') then   //���� � ���������� ����� ������ �� �������...
      processingMassMszFact();


    //��������� ������ � ��������
    renderTableMszFact();

    //��������� ���������� � ���� CSV
    if (M_Error.Text = '') then   //���� � ���������� ����� ������ �� �������...
      renderDataToCsv();


    //����� ���������
    if (M_Info.Text <> '') then
      Begin
        RE_Log.Lines.Add('');
        RE_Log.Lines.Add('������������ �� ����������� ������:');
        RE_Log.Text := RE_Log.Text + M_Info.Text;
      End;

    //����� ������
    if (M_Error.Text <> '') then
      Begin
        RE_Log.Lines.Add('');
        RE_Log.Lines.Add('������������ ����������� ������:');
        RE_Log.Text := RE_Log.Text + M_Error.Text;
      End;

    B_log.Click; //��������� ����� � ������ ���� ��� �� ���� �������

  End;

end;

procedure TF_main.FormCreate(Sender: TObject);
begin
  //��������� �������������� ���������
    //���������� ��������� ����
    F_main.Caption := '��������� ������ ���������� ��� xlsx->csv.  ������ ��������� ' + VersionSoftware + '  ������ ������� ' + VersionTemplate;

    //���������� ������� "� ���������"
    L_VersionSoftware.caption := VersionSoftware;
    L_VersionTemplate.Caption := VersionTemplate;
    L_DateBuild.Caption := DateBuild;

    //������� �������� � RE_log
    RE_Log.Paragraph.FirstIndent:=10;
    RE_Log.Paragraph.LeftIndent:=10;

    //�������� ��������� �� �������� �����
    B_log.Click;

end;

end.
