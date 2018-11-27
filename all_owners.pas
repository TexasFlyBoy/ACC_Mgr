unit all_owners;

interface

uses
  vcl.dbgrids, Classes, Graphics, Dialogs, Main, SysUtils;

function AllOwnersOpen: Boolean;
function AllOwnersColumnSetup(Setup: boolean): Boolean;

implementation

{ ----------------------------------------------------------------------+
  DBGrid2CellClick():
  This procedure finds the houseAcct value of the current record
  in the All_Owners grid and copies it to the "Acct Search" TEdit
  box. It then calls the eCurrentAcctSearchChange event.
  +----------------------------------------------------------------------- }
procedure AllOwnersCellClick(Column: TColumn);
begin
  MainForm.eCurrentAcctSearch.Text := MainForm.adoTableOwners.FieldValues['houseAcct'];
  MainForm.eCurrentAcctSearchChange(Column);
end;

function AllOwnersOpen: Boolean;
begin
  AllOwnersOpen := False;
  try
    MainForm.adoTableOwners.Active := True;
    AllOwnersOpen := True;
  except
    MainForm.adoTableOwners.Active := False;
    AllOwnersOpen := False;
  end;
end;

{-----------------------------------------------------------------------------+
   AllOwnersColumnSetup(Setup: boolean):

   This function deletes any existing columns in the dbgridAllOwners.
   If the argument is TRUE then it creates the columns using a
   hardcoded methodology.
   This is done to reduce the file size of the executable.
 +----------------------------------------------------------------------------}
function AllOwnersColumnSetup(Setup: boolean): Boolean;
var
  i: Integer;
begin
  with MainForm.dbgridAllOwners do begin

    if Columns.Count > 0 then
    try
      for I := (Columns.Count - 1) downto 0 do begin
        Columns.Delete(i);
      end;
    except
      // do something here
    end;

    if Not(Setup) then exit;

    for i := 0 to 16 do begin
      Columns.Add;
      Columns.Items[i].Visible := True;
      Columns.Items[i].Title.Color := $00F0B4DC;
      Columns.Items[i].Title.Font.Style := [fsBold];
      Columns.Items[i].Alignment := taLeftJustify;
      Columns.Items[i].Width := 80;
    end;

    Columns.Items[0].FieldName := 'ownerID';
    Columns.Items[0].Title.Caption := 'ID #';
    Columns.Items[0].Alignment := taRightJustify;

    Columns.Items[1].FieldName := 'houseAcct';
    Columns.Items[1].Title.Caption := 'Acct';
    Columns.Items[1].Alignment := taRightJustify;

    Columns.Items[2].FieldName := 'Owner';
    Columns.Items[2].Title.Caption := Columns.Items[2].FieldName;
    Columns.Items[2].Width := 200;

    Columns.Items[3].FieldName := 'Offsite';
    Columns.Items[3].Title.Caption := Columns.Items[3].FieldName;
    Columns.Items[3].Width := 100;

    Columns.Items[4].FieldName := 'Phone';
    Columns.Items[4].Title.Caption := Columns.Items[4].FieldName;

    Columns.Items[5].FieldName := 'Alt Phone';
    Columns.Items[5].Title.Caption := Columns.Items[5].FieldName;

    Columns.Items[6].FieldName := 'Mobile 1';
    Columns.Items[6].Title.Caption := Columns.Items[6].FieldName;

    Columns.Items[7].FieldName := 'Mobile 2';
    Columns.Items[7].Title.Caption := Columns.Items[7].FieldName;

    Columns.Items[8].FieldName := 'closeDate';
    Columns.Items[8].Title.Caption := 'Close Date';
    Columns.Items[8].Width := 70;

    Columns.Items[9].FieldName := 'sellDate';
    Columns.Items[9].Title.Caption := 'Sell Date';
    Columns.Items[9].Width := 70;

    Columns.Items[10].FieldName := 'greeting';
    Columns.Items[10].Title.Caption := 'Greeting';
    Columns.Items[10].Width := 120;

    Columns.Items[11].FieldName := 'mailName';
    Columns.Items[11].Title.Caption := 'Mail Name';
    Columns.Items[11].Width := 150;

    Columns.Items[12].FieldName := 'Address';
    Columns.Items[12].Title.Caption := 'Offsite Address';
    Columns.Items[12].Width := 120;

    Columns.Items[13].FieldName := 'city';
    Columns.Items[13].Title.Caption := 'Offsite City';
    Columns.Items[13].Width := 70;

    Columns.Items[14].FieldName := 'state';
    Columns.Items[14].Title.Caption := 'Offsite State';

    Columns.Items[15].FieldName := 'Zip';
    Columns.Items[15].Title.Caption := 'Offsite Zip';
    Columns.Items[15].Alignment := taRightJustify;

    Columns.Items[16].FieldName := 'cityStateZip';
    Columns.Items[16].Visible := False;
  end;
end;




end.
