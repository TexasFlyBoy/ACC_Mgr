unit ini_file;

interface

uses SysUtils, Windows, Messages, Classes, Graphics, Controls, Inifiles,
      Main,Vcl.Forms, Dialogs;

procedure CreateIniFile;
function IniFileExists: Boolean;
function ReadIniString(section, key: string): string overload;
function ReadIniBoolean(section, key: string): Boolean overload;
function ReadIniColor(section, key: string): TColor overload;
function ReadIniInteger(section, key: string): Integer overload;
function ReadIniDouble(section, key: string): double overload;
function WriteIniString(section, key, value: string): Boolean overload;
function WriteIniBoolean(section, key: string; value: Boolean): Boolean overload;
function WriteIniColor(section, key: string; value: TColor): Boolean overload;
function WriteIniInteger(section, key: string; value: Integer): Boolean overload;
function WriteIniDouble(section, key: string; value: double): Boolean overload;
function IniSqlDirectoryExists(section, key: string): string;

implementation


procedure CreateIniFile;
var
  Ini: TIniFile;
  fileName : string;
  myFile   : TextFile;
  data     : string;
begin
  fileName := ChangeFileExt(Application.ExeName, '.INI' );
  if FileExists(fileName) then Exit;
  Ini := TIniFile.Create(fileName);
  try
    Ini.WriteString( 'Connection', 'ConnString', '44');
    Ini.WriteInteger( 'Form', 'Left', 50);
    Ini.WriteString( 'Form', 'Caption', 'what is going on here' );
    Ini.WriteFloat('Colors\Panels','pnlCurrentOwners',     $0084B984);
    Ini.WriteFloat('Colors\Panels','pnlAcAppEnter',        $0061AE0C);
    Ini.WriteFloat('Colors\Panels','pnlVioStatus',         $00F1A65A);
    Ini.WriteFloat('Colors\Panels','pnlGenVioLetterEnter', $004798E0);
    Ini.WriteFloat('Colors\Panels','pnlHousesEnter',       $00DE8B3E);
    Ini.WriteFloat('Colors\Panels','pnlOwnersEnter',       $00399AE1);
    Ini.WriteFloat('Colors\Panels','pnlWelcomeEnter',      $0054C622);
    Ini.WriteFloat('Colors\Panels','pnlAllGenVioLetters',  $004998E0);
    Ini.WriteFloat('Colors\Panels','pnlAcApps1',           $00D1B499);
    Ini.WriteFloat('Colors\Panels','pnlAcApps2',           $00D1B499);
    Ini.WriteFloat('Colors\Panels','pnlGenLetters1',       $00EEC29B);
    Ini.WriteFloat('Colors\Panels','pnlGenLetters2',       $00EEC29B);
    Ini.WriteFloat('Colors\Panels','pnlMemoToLegal',       $00FB5942);
    Ini.WriteFloat('Colors\Panels','pnlSqlButtons',        $00FF901E);

    Ini.WriteFloat('Colors','dbGridCurrentOwners',         $00B1D2B0);
    Ini.WriteFloat('Colors','AC Applications Grid',        $0097F231);
    Ini.WriteFloat('Colors','Violations Grid',             $00EFD0C9);
    Ini.WriteFloat('Colors','Current Owners Grid',         $00B1D2B0);
    Ini.WriteFloat('Colors','Current Owners Grid',         $00B1D2B0);
    Ini.WriteFloat('Colors','dbGridSqlView',               $00FACE87);

    Ini.WriteString('VioLetters\Defaults','letterType', 'GEN');
    Ini.WriteString('VioLetters\Defaults','letterNumber', 'FOL');
    Ini.WriteString('VioLetters\Defaults','specificViolationTitle', 'Specific Violation Title');
    Ini.WriteString('VioLetters\Defaults','textViolationWords', 'Text Violation Words');
    Ini.WriteString('VioLetters\Defaults','remedyWords', 'Remedy Words');
    Ini.WriteString('VioLetters\Defaults','senderInitials', 'Bozo');
    Ini.WriteString('VioLetters\Defaults','3OLSignLine', 'Clown, Bozo D.');
    Ini.WriteString('VioLetters\Defaults','4OLSignLine', 'Bozo D. Clown');
    Ini.WriteString('VioLetters\Defaults','209SignLine', 'Bozo Dee Clown');

    Ini.WriteString('sql\Directory','AutoFormat', '');

  finally
    Ini.Free;
  end;
end;

{--------------------------------------------------------------------------+
|  function IniFileExists: boolean;                                        |
|  Determines of the iniFile for the application exists in the directory   |
|  of the application.                                                     |
+--------------------------------------------------------------------------}
function IniFileExists: boolean;
var
  filename: string;
begin
  fileName := ChangeFileExt( Application.ExeName, '.INI' );
  IniFileExists := FileExists(fileName);
end;

{--------------------------------------------------------------------------+
|  function ReadIniString(section, key: string): string overload;          |
|  Read the value of the key from the section of the iniFile               |
+--------------------------------------------------------------------------}
function ReadIniString(section, key: string): string overload;
var
  Ini: TIniFile;
  fileName, value: string;
begin
  ReadIniString := '';
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      ReadIniString := Ini.ReadString(section, key, '');
    finally
      Ini.Free
    end;
  end;
end;

{--------------------------------------------------------------------------+
|  function ReadIniBoolean(section, key: string): boolean overload;        |
|  Read the value of the key from the section of the iniFile               |
+--------------------------------------------------------------------------}
function ReadIniBoolean(section, key: string): boolean overload;
var
  Ini: TIniFile;
  fileName: string;
begin
  ReadIniBoolean := False;
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      ReadIniBoolean := Ini.ReadBool(section, key, False);
    finally
      Ini.Free
    end;
  end;
end;

{--------------------------------------------------------------------------+
|  function ReadIniColor(section, key: string): TColor overload;           |
|  Read the value of the key from the section of the iniFile               |
+--------------------------------------------------------------------------}
function ReadIniColor(section, key: string): TColor overload;
var
  Ini: TIniFile;
  fileName: string;
begin
  ReadIniColor := 0;
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      ReadIniColor := Ini.ReadInteger(section, key, 0);
    finally
      Ini.Free
    end;
  end;
end;

{--------------------------------------------------------------------------+
|  function ReadIniInteger(section, key: string): integer overload;        |
|  Read the value of the key from the section of the iniFile               |
+--------------------------------------------------------------------------}
function ReadIniInteger(section, key: string): integer overload;
begin
  ReadIniInteger := ReadIniColor(section, key);
end;

{--------------------------------------------------------------------------+
|  function ReadIniDouble(section, key: string): double overload;          |
|  Read the value of the key from the section of the iniFile               |
+--------------------------------------------------------------------------}
function ReadIniDouble(section, key: string): double overload;
var
  Ini: TIniFile;
  fileName: string;
begin
  ReadIniDouble := 0.0;
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      ReadIniDouble := Ini.ReadFloat(section, key, 0.0);
    finally
      Ini.Free
    end;
  end;
end;

{------------------------------------------------------------------------------------+
|  function WriteIniString(section, key, value: string): boolean overload;           |
|  Write the value of the key to the section of the iniFile                          |
+------------------------------------------------------------------------------------}
function WriteIniString(section, key, value: string): boolean overload;
var
  Ini: TIniFile;
  fileName: string;
begin
  WriteIniString := False;
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if not(FileExists(fileName)) then
    CreateIniFile;
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      Ini.WriteString(section, key, value);
      WriteIniString := True;
    finally
      Ini.Free
    end;
  end;
end;

{------------------------------------------------------------------------------------+
|  function WriteIniBoolean(section, key: string; value: boolean): boolean overload; |
|  Write the value of the key to the section of the iniFile                          |
+------------------------------------------------------------------------------------}
function WriteIniBoolean(section, key: string; value: boolean): boolean overload;
var
  Ini: TIniFile;
  fileName: string;
begin
  WriteIniBoolean := False;
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      Ini.WriteBool(section, key, value);
      WriteIniBoolean := True;
    finally
      Ini.Free
    end;
  end;
end;

{---------------------------------------------------------------------------------+
|  function WriteIniColor(section, key: string; value: TColor): boolean overload; |
|  Write the value of the key to the section of the iniFile                       |
+---------------------------------------------------------------------------------}
function WriteIniColor(section, key: string; value: TColor): boolean overload;
var
  Ini: TIniFile;
  fileName: string;
begin
  WriteIniColor := False;
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      Ini.WriteInteger(section, key, value);
      WriteIniColor := True;
    finally
      Ini.Free
    end;
  end;
end;

{------------------------------------------------------------------------------------+
|  function WriteIniInteger(section, key: string; value: integer): boolean overload; |
|  Write the value of the key to the section of the iniFile                          |
+------------------------------------------------------------------------------------}
function WriteIniInteger(section, key: string; value: integer): boolean overload;
begin
  WriteIniInteger := WriteIniColor(section, key, value);
end;

{----------------------------------------------------------------------------------+
|  function WriteIniDouble(section, key: string; value: double): boolean overload; |
|  Write the value of the key to the section of the iniFile                        |
+----------------------------------------------------------------------------------}
function WriteIniDouble(section, key: string; value: double): boolean overload;
var
  Ini: TIniFile;
  fileName: string;
begin
  WriteIniDouble := False;
  filename := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(fileName) then begin
    try
      Ini := TIniFile.Create(fileName);
      Ini.WriteFloat(section, key, value);
      WriteIniDouble := True;
    finally
      Ini.Free
    end;
  end;
end;

function IniSqlDirectoryExists(section, key: string): string;
var
  Ini: TIniFile;
  myFileName, sqlDirName: string;
begin
  if not(IniFileExists) then
    CreateIniFile;

  myFileName := ChangeFileExt(Application.ExeName, '.ini');
  if FileExists(myFileName) then begin
    try
      Ini := TIniFile.Create(myFileName);
      sqlDirName := Ini.ReadString(section, key, '');
      IniSqlDirectoryExists := sqlDirName + '\\';
      if not(DirectoryExists(sqlDirName)) then begin
        IniSqlDirectoryExists := '';
        with TFileOpenDialog.Create(nil) do
          try
            Options := [fdoPickFolders];
            if Execute then
              sqlDirName := FileName + '\\';
          finally
            Free;
          end;
        Ini.WriteString(section, key, sqlDirName);
        IniSqlDirectoryExists := sqlDirName;
      end;
    finally
      Ini.Free
    end;
  end;
end;


end.
