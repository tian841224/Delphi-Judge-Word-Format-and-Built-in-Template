unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,StdCtrls,ComObj,TlHelp32,FileCtrl, ExtCtrls, Buttons;

type
  TCheckTest = class(TForm)
    dlgOpen1: TOpenDialog;
    pnl1: TPanel;
    lblMessage: TLabel;
    btnOutput: TButton;
    mmoMessage: TMemo;
    btninput: TButton;
    btnClose: TButton;
    procedure btninputClick(Sender: TObject);
    procedure btnOutputClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    function KillWordTask : integer;
    procedure Comparison (TitleTemp :TStringList);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btn1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  CheckTest: TCheckTest;

implementation

{$R *.dfm}

procedure TCheckTest.btn1Click(Sender: TObject);

begin
  lblMessage.Font.Name := 'Noto Sans TC Regular';
  btninput.Font.Name := 'Noto Sans TC Regular';
  btnOutput.Font.Name := 'Noto Sans TC Regular';
  btnClose.Font.Name := 'Noto Sans TC Regular';
end;

procedure TCheckTest.btnCloseClick(Sender: TObject);
begin
Close;
end;

procedure TCheckTest.btninputClick(Sender: TObject);
var
WordApp,WordDoc, myRange : Variant ;
i,x,y : Integer;
TitleTemp  : TStringList;
check : Boolean;
begin
  check := False;
  KillWordTask;
  TitleTemp := TStringList.Create ;
  WordApp := CreateOleObject('Word.Application');
  if dlgOpen1.Execute then  WordDoc := WordApp.Documents.Open(dlgOpen1.FileName)
  else
    begin
      ShowMessage('請選擇檔案');
      Exit;
    end;
  myRange := WordDoc.Content;
try
  mmoMessage.Clear;
  try
    if WordDoc.Tables.Item(1).Columns.Count  <> 26 then mmoMessage.Lines.Add('Word行數不正確');
  except
    mmoMessage.Lines.Add('Word行數不正確');
    Exit;
  end;

  mmoMessage.Lines.Add('－－－－－－比對資料中...－－－－－－');

  for i  := 2 to WordDoc.Tables.Item(1).Columns.Count   do
    begin
      mmoMessage.Lines.Add('－－－－－－目前進度'+IntToStr(i)+'/'+IntToStr(WordDoc.Tables.Item(1).Columns.Count)+'－－－－－－');
      if (i = 12)  then
        begin
          for x  := 2 to WordDoc.Tables.Item(1).Rows.Count  do
            begin
              for y:=1 to length(Trim(WordDoc.Tables.Item(1).Cell(x,i).Range.Text)) do
                if not (Trim(WordDoc.Tables.Item(1).Cell(x,i).Range.Text)[y] in ['0'..'9','.','-']) and (check = False)  then
                  begin
                    mmoMessage.Lines.Add('設定錯誤：[ 第'+IntToStr(i)+'列、第'+IntToStr(x)+'行 ] 須為數字');
                    check := True;
                  end;
                  check := False;
            end;
        end

      else if (i = 2 ) or (i = 8 ) or (i =9) or (i =10) or (i =11) or (i =13) or (i =14) or (i =15) or (i =17) or (i =19) then
        begin
          for x  := 2 to WordDoc.Tables.Item(1).Rows.Count  do
            begin
              if  Trim(WordDoc.Tables.Item(1).Cell(x,i).Range.Text) = '' then
                mmoMessage.Lines.Add('設定錯誤：[ 第'+IntToStr(i)+'列、第'+IntToStr(x)+'行 ] 不能為空')
            end;
        end ;
      TitleTemp.Add(Trim(WordDoc .Tables.Item(1).Cell(1,i).Range.Text));
    end;
  Comparison(TitleTemp);
finally
  mmoMessage.Lines.Add('－－－－－－比對結束－－－－－－');
  WordApp.Quit;
  WordApp:=Unassigned;
end;
end;

procedure TCheckTest.btnOutputClick(Sender: TObject);
Var
RCS : TResourceStream;
dirpath : string;
begin
mmoMessage.Clear;
if SelectDirectory('選擇匯出位置','',dirpath) then
  begin
    RCS := TResourceStream.Create(HInstance, 'WordExample','Word');
    RCS.SaveToFile(dirpath+'\QuExp.docx'); //另存檔案
    RCS.Free;
  end;
  mmoMessage.Text := '---匯出成功---';
end;
function TCheckTest.KillWordTask : integer;
const
  PROCESS_TERMINATE=$0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  result := 0;

  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);

  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = 'WINWORD.EXE') or
       (UpperCase(FProcessEntry32.szExeFile) = 'WINWORD.EXE')) then
      Result := Integer(TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0),FProcessEntry32.th32ProcessID), 0));
    ContinueLoop := Process32Next(FSnapshotHandle,FProcessEntry32);
  end;

  CloseHandle(FSnapshotHandle);
end;

procedure TCheckTest.Comparison (TitleTemp :TStringList);
var
TitleArray  : TStringList;
i : Integer;
begin
  TitleArray := TStringList.Create;
  TitleArray.Add('題目');
  TitleArray.Add('1選項');
  TitleArray.Add('2選項');
  TitleArray.Add('3選項');
  TitleArray.Add('4選項');
  TitleArray.Add('5選項');
  TitleArray.Add('標準答案');
  TitleArray.Add('題型');
  TitleArray.Add('類別');
  TitleArray.Add('考科');
  TitleArray.Add('年度');
  TitleArray.Add('題序');
  TitleArray.Add('程度');
  TitleArray.Add('科目');
  TitleArray.Add('知識點');
  TitleArray.Add('難易度');
  TitleArray.Add('詳解');
  TitleArray.Add('命題者');
  TitleArray.Add('審題者');
  TitleArray.Add('題組題');
  TitleArray.Add('爭議題');
  TitleArray.Add('QID');
  TitleArray.Add('答題次');
  TitleArray.Add('正答率');
  TitleArray.Add('類似題');
  mmoMessage.Lines.Add('－－－－－－比對標題－－－－－－');
  for i := 0 to TitleArray.Count -1 do
    if TitleArray[i] <> Trim(TitleTemp[i])  then  mmoMessage.Lines.Add('設定錯誤：[ 第'+IntToStr(i)+'列 ] 標題錯誤');
end;

procedure TCheckTest.FormDestroy(Sender: TObject);
begin
  RemoveFontResource(PChar('NotoSansTC.otf'));
  SendMessage(HWND_BROADCAST,WM_FONTCHANGE,0,0);
end;

procedure TCheckTest.FormShow(Sender: TObject);
Var
RCS : TResourceStream;
begin
  RCS := TResourceStream.Create(hInstance, 'NotoSansTC', Pchar('otf'));
  RCS.SavetoFile('NotoSansTC.otf');
  RCS.Free;
  AddFontResource(PChar('NotoSansTC.otf'));
  SendMessage(HWND_BROADCAST,WM_FONTCHANGE,0,0);

  lblMessage.Font.Name := 'Noto Sans TC Regular';
  btninput.Font.Name := 'Noto Sans TC Regular';
  btnOutput.Font.Name := 'Noto Sans TC Regular';
  btnClose.Font.Name := 'Noto Sans TC Regular';
end;


end.
