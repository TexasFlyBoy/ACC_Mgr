unit genVioLetters;

interface

uses
  vcl.dbgrids, Classes, Graphics, Dialogs, Main;

function genVioLettersOpen: Boolean;


implementation

function genVioLettersOpen: Boolean;
begin
  genVioLettersOpen := False;
  try
    MainForm.AdoDataSetVioLetters.Active := True;
    genVioLettersOpen := True;
  except
    MainForm.AdoDataSetVioLetters.Active := False;
    genVioLettersOpen := False;
  end;
end;


end.
