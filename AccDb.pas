unit AccDb;

interface

uses
  vcl.dbgrids, Classes, Graphics, Dialogs, Main, SysUtils;

  function MakeConnection(Connect: boolean): boolean;

implementation

uses
  all_owners, all_letters;

function MakeConnection(Connect: boolean): boolean;
begin
  MakeConnection := False;
  try
    with MainForm do begin
      sbCurrentOwners.Panels[2].Text := ADOConnection1.ConnectionString;
      AdoTableGenVioLetters.Active := Connect;
      AdoDataSetOwners.Active := Connect;
      adoTblHouses.Active := Connect;
      adoTableOwners.Active := Connect;
      if (AllOwnersOpen) then
        AllOwnersColumnSetup(Connect);
      AllLettersColumnSetup;

      adoTblAllApprovalLetters.Active := Connect;
//      adoTblOffsiteOwners.Active := Connect;

      adoTblWelcomeLetters.Active := Connect;
      AdoTableCurrentOwners.Active := Connect;
      adoTblBrowseGenVioLetters.Active := Connect;
      adoTblMemoToLegal.Active := Connect;
      adoTblPropInLegal.Active := Connect;
      adoTblLegalStatus.Active := Connect;

      MakeConnection := True;
    end;
  except
    ShowMessage('There was a problem connecting the tables.');
  end;
end;




end.
