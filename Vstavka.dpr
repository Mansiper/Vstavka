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
  Result:='ХУ';
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
        Result:='ХУ';
        Exit;
      end;
    end;
end;

//------------------------------------------------------------------------------
  //Не допустить запуск второй копии
function Vyhod: Boolean;
begin
  Handle:=OpenMutex(MUTEX_ALL_ACCESS, False, 'MutVst');
  Result:=(Handle<>0);
  If Handle=0 then
    Handle:=CreateMutex(nil, False, 'MutVst');
end;

//------------------------------------------------------------------------------

begin
  If Vyhod then Exit; //Дабы не допустить запуск второй копии программы

  Handle:=GetForegroundWindow;  //Получаю хэндл проги
  ShowWindow(Handle, SW_HIDE);  //Скрытие окна

  CoInitialize(nil);  //Инициализация для консольного приложения
                      //при использовании COM-объектов
  //Назначение символов-исключений
  Nizya:=['\', '|', ':', '*', '?', '"', '<', '>', '/'];

  M:=TMemo.CreateParented(Handle);  //Создание компонента типа TMemo

  While True do
  Begin
    Sleep(2000);
//--------------
    Handle:=GetForegroundWindow;      //Получение хэндла активного окна
    Put:=GetAddressByHandle(Handle);  //Получение адреса страницы
//--------------
    //Изменение написания адреса
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
    if ClipBoard.HasFormat(cf_text) then  //Если в буфере текст
    begin
      M.Lines.Text:=ClipBoard.AsText; //Получение текста из буфера
      nom:=0;
      repeat
        str:=M.Lines.Strings[nom];  //Первая строка - имя файла
        str:=Trim(str);             //Удаление пробелов в начале и в конце
        Inc(nom);
      until (Length(str)<>0);

      nom:=1;
      repeat                        //Удаление символов-исключений
        if str[nom] in Nizya then
          Delete(str, nom, 1)
        else Inc(nom);
      until nom=Length(str)+1;
      str:=Trim(str);               //Ещё раз удаляю пробелы
    end
    else Continue;
//--------------
    //Создание/изменение файла
    if not FileExists(Put+'\'+str+'.txt') then
    begin
      AssignFile(F, Put+'\'+str+'.txt');
      Rewrite(F);
      Write(F, M.Lines.Text);
      CloseFile(F);
    end//Если файла нет

    else if MessageDlg('Файл уже существует. Переписать?', mtConfirmation,
    [mbYes, mbNo], 0)=IDYES then
    begin
      AssignFile(F, Put+'\'+str+'.txt');
      Rewrite(F);
      Write(F, M.Lines.Text);
      CloseFile(F);
    end;//Если файл есть
    Sleep(200);
  End;
//--------------
  M.Destroy;
  CoUninitialize();                 //Завершение работы с COM-объектами
end.

//  FreeConsole;  //Скрытие окна
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

{ //В проге
    Hook:=SetWindowsHookEx(13(*WH_KEYBOARD_LL*), LowLevelMouseProc,
     GetModuleHandle(nil), 0);} //  Hook: HHOOK;
