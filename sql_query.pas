unit sql_query;

interface

uses
  SysUtils, Classes, Graphics;

procedure btnShowSqlClick1(Sender: TObject);
procedure btnRunSqlClick;

implementation

uses
  Main;


procedure btnShowSqlClick1(Sender: TObject);
var
  i: Integer;
begin
  mainForm.memoSqlText.Lines.Clear;
  for I := 0 to (mainform.adoQryCullLetters.SQL.Count - 1) do
  begin
  mainForm.memoSqlText.Lines.Insert(i, mainForm.adoQryCullLetters.SQL.Strings[i]);
  end;
end;

procedure btnRunSqlClick;
var
  i: Integer;
begin
  with mainForm.adoQryCullLetters do begin
    Active := False;
    SQL.Clear;
    // Delete the line if it's a comment or a blank line
    for I := (mainForm.memoSqlText.Lines.Count - 1) downto 0 do
      if (Trim(mainForm.memoSqlText.Lines.Strings[i]).StartsWith('#')) then
        mainForm.memoSqlText.Lines.Delete(i)
      else
      if (Length(Trim(mainForm.memoSqlText.Lines.Strings[i])) < 1) then
        mainForm.memoSqlText.Lines.Delete(i);

    // Now insert the lines
    for I := 0 to (mainForm.memoSqlText.Lines.Count - 1) do
        SQL.Insert(i, mainForm.memoSqlText.Lines.Strings[i]);
    Active := True;
    mainForm.memoSqlText.Lines.Append('');
    mainForm.memoSqlText.Lines.Append('');
    mainForm.memoSqlText.Lines.Append('#  SQL Comments Begin With Hash Symbol');
    mainForm.memoSqlText.Lines.Append('#  SQL Wildcard Character: %');
    mainForm.memoSqlText.Lines.Append('#  Here are the field names in the query result');
    mainForm.memoSqlText.Lines.Append('');

  end;

  for I := 0 to mainForm.dbGridSqlView.Columns.Count - 1 do
    begin
      mainForm.dbGridSqlView.Columns[i].Width := 100;
      mainForm.dbGridSqlView.Columns[i].Title.Font.Style := [fsBold];
      mainForm.dbGridSqlView.Columns[i].Title.Alignment := taCenter;
      mainForm.memoSqlText.Lines.Append('#  ' + mainForm.dbGridSqlView.Columns[i].FieldName + '  ,');
      MainForm.dbGridSqlView.FixedColor := $FACE87;   // Light Sky Blue
    end;
  mainForm.sbSqlView.Panels[1].Text := 'Num of Records: ' + IntToStr(mainForm.adoQryCullLetters.RecordCount);
  mainForm.sbSqlView.Panels[2].Text := 'Num of Fields: ' + IntToStr(mainForm.adoQryCullLetters.FieldCount);

end;


end.
