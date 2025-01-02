program Vstavka;

{$APPTYPE CONSOLE}

uses
  Windows, ShDocVw, ActiveX, Classes, Dialogs, SysUtils, StdCtrls, Clipbrd;

var
  Nizya: Set Of Char;
  str: ShortString;
  Nom: Word;
  Handle: HWND;
  Put: String;
  F: TextFile;
  M: TMemo;

//------------------------------------------------------------------------------

function GetAddressByHandle(Handle: Integer): String;
var ShellWindows: IShellWindows;
    i: Integer;
begin
  Result:='��';
  ShellWindows:=CoShellWindows.Create;
  For i:=0 to ShellWindows.Count-1 do
    try
      if (ShellWindows.Item(i) as IWebBrowser2Disp).HWND=Handle then
      begin
        Result:=(ShellWindows.Item(i) as IWebBrowser2Disp).LocationURL;
        Break;
      end;
    except
      on Exception do
      begin
        Result:='��';
        Exit;
      end;
    end;
end;

//------------------------------------------------------------------------------
  //�� ��������� ������ ������ �����
function Vyhod: Boolean;
begin
  Handle:=OpenMutex(MUTEX_ALL_ACCESS, False, 'MutVst');
  Result:=(Handle<>0);
  If Handle=0 then
    Handle:=CreateMutex(nil, False, 'MutVst');
end;

//------------------------------------------------------------------------------

begin
  If Vyhod then Exit; //���� �� ��������� ������ ������ ����� ���������

  Handle:=GetForegroundWindow;  //������� ����� �����
  ShowWindow(Handle, SW_HIDE);  //������� ����

  CoInitialize(nil);  //������������� ��� ����������� ����������
                      //��� ������������� COM-��������
  //���������� ��������-����������
  Nizya:=['\', '|', ':', '*', '?', '"', '<', '>', '/'];

  M:=TMemo.CreateParented(Handle);  //�������� ���������� ���� TMemo

  While True do
  Begin
    Sleep(2000);
//--------------
    Handle:=GetForegroundWindow;      //��������� ������ ��������� ����
    Put:=GetAddressByHandle(Handle);  //��������� ������ ��������
//--------------
    //��������� ��������� ������
    if Pos('file', Put)=1 then
    begin
      Delete(Put, 1, 8);
      while True do
      begin
        Nom:=Pos('%20', Put);
        if Nom=0 then Break;
        Delete(Put, Nom, 3);
        Insert(' ', Put, Nom);
      end;
    end
    else Continue;

    if (GetAsyncKeyState(VK_CONTROL)=0) or
       (GetAsyncKeyState(Ord('V'))=0) or
       (GetAsyncKeyState(VK_MENU)<>0) or
       (GetAsyncKeyState(VK_SHIFT)<>0) then Continue;

//--------------
    if ClipBoard.HasFormat(cf_text) then  //���� � ������ �����
    begin
      M.Lines.Text:=ClipBoard.AsText; //��������� ������ �� ������
      nom:=0;
      repeat
        str:=M.Lines.Strings[nom];  //������ ������ - ��� �����
        str:=Trim(str);             //�������� �������� � ������ � � �����
        Inc(nom);
      until (Length(str)<>0);

      nom:=1;
      repeat                        //�������� ��������-����������
        if str[nom] in Nizya then
          Delete(str, nom, 1)
        else Inc(nom);
      until nom=Length(str)+1;
      str:=Trim(str);               //��� ��� ������ �������
    end
    else Continue;
//--------------
    //��������/��������� �����
    if not FileExists(Put+'\'+str+'.txt') then
    begin
      AssignFile(F, Put+'\'+str+'.txt');
      Rewrite(F);
      Write(F, M.Lines.Text);
      CloseFile(F);
    end//���� ����� ���

    else if MessageDlg('���� ��� ����������. ����������?', mtConfirmation,
    [mbYes, mbNo], 0)=IDYES then
    begin
      AssignFile(F, Put+'\'+str+'.txt');
      Rewrite(F);
      Write(F, M.Lines.Text);
      CloseFile(F);
    end;//���� ���� ����
    Sleep(200);
  End;
//--------------
  M.Destroy;
  CoUninitialize();                 //���������� ������ � COM-���������
end.

//  FreeConsole;  //������� ����
//  Handle:=GetStdHandle(STD_OUTPUT_HANDLE);

//------------------------------------------------------------------------------

{function LowLevelMouseProc(
     nCode:   LongInt;  // hook code
     wParam:  WPARAM;   // message identifier
     lParam:  LPARAM    // message data
    ):  LRESULT; stdcall;
begin
//-----
end;}

{ //� �����
    Hook:=SetWindowsHookEx(13(*WH_KEYBOARD_LL*), LowLevelMouseProc,
     GetModuleHandle(nil), 0);} //  Hook: HHOOK;
