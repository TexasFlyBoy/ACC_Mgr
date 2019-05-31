unit all_letters;

interface


uses ADODB,
// vcl.DBGrids,
Classes,
Graphics,
Dialogs,
SysUtils;


function AllLettersOpen: Boolean;
function AllLettersColumnSetup: Boolean;


implementation

uses Main;

function AllLettersOpen: Boolean;
begin
  AllLettersOpen := False;
  try
    MainForm.adoTblAllLetters.Active := True;
    AllLettersOpen := True;
  except
    MainForm.adoTblAllLetters.Active := False;
    AllLettersOpen := False;
  end;
end;

function AllLettersColumnSetup: Boolean;
var
  i, numOfColumns, numOfVisCol: Integer;
begin
  AllLettersColumnSetup := False;
  numOfColumns := 42;
  numOfVisCol := 12;
  with MainForm.dbGridAllLetters do begin

      if Columns.Count > 0 then
        for i := (Columns.Count - 1) downto 0 do
          Columns.Delete(i);

      for i := 0 to (numOfColumns - 1) do begin
        Columns.Add;
        if (i < (numOfVisCol - 1)) then
          Columns.Items[i].Visible := True
        else
          Columns.Items[i].Visible := False;
        Columns.Items[i].Title.Color := $00F8E3D1;
        Columns.Items[i].Title.Font.Style := [fsBold];
        Columns.Items[i].Title.Alignment := taCenter;
        Columns.Items[i].Alignment := taLeftJustify;
        Columns.Items[i].Width := 50;
      end;
      // These are the visible fields
      Columns.Items[0].FieldName := 'houseAcct';
      Columns.Items[0].Title.Caption := 'Acct';
      Columns.Items[0].Width := 50;

      Columns.Items[1].FieldName := 'violationId';
      Columns.Items[1].Alignment := taCenter;
      Columns.Items[1].Title.Caption := 'Vio #';
      Columns.Items[1].Width := 55;

      Columns.Items[2].FieldName := 'letterNumber';
      Columns.Items[2].Alignment := taCenter;
      Columns.Items[2].Title.Caption := 'Letter #';
      Columns.Items[2].Width := 60;

      Columns.Items[3].FieldName := 'letterType';
      Columns.Items[3].Alignment := taCenter;
      Columns.Items[3].Title.Caption := 'Type';
      Columns.Items[3].Width := 55;

      Columns.Items[4].FieldName := 'permitNumber';
      Columns.Items[4].Alignment := taCenter;
      Columns.Items[4].Title.Caption := 'Permit #';

      Columns.Items[5].FieldName := 'owner';
      Columns.Items[5].Title.Caption := 'Owner';
      Columns.Items[5].Width := 205;

      Columns.Items[6].FieldName := 'address';
      Columns.Items[6].Title.Caption := 'Address';
      Columns.Items[6].Width := 125;

      Columns.Items[7].FieldName := 'letterDate';
      Columns.Items[7].Title.Caption := 'Letter Date';
      Columns.Items[7].Width := 70;

      Columns.Items[8].FieldName := 'senderInitials';
      Columns.Items[8].Alignment := taCenter;
      Columns.Items[8].Title.Caption := 'Initials';
      Columns.Items[8].Width := 60;

      Columns.Items[9].FieldName := 'offsite';
      Columns.Items[9].Width := 100;

      Columns.Items[10].FieldName := 'presidentLine';
      Columns.Items[10].Title.Caption := 'Signature Line';
      Columns.Items[10].Width := 200;

      // These are the hidden fields
      Columns.Items[11].FieldName := 'ccField';
      Columns.Items[12].FieldName := 'specificApprovalWords';
      Columns.Items[13].FieldName := 'permitDateIssued';
      Columns.Items[14].FieldName := 'permitMonth';
      Columns.Items[15].FieldName := 'permitYear';
      Columns.Items[16].FieldName := 'applicationDate';
      Columns.Items[17].FieldName := 'approvalDate';
      Columns.Items[18].FieldName := 'permitLine';
      Columns.Items[19].FieldName := 'greeting';
      Columns.Items[20].FieldName := 'offsiteCity';
      Columns.Items[21].FieldName := 'firstLine';
      Columns.Items[22].FieldName := 'ownerId';
      Columns.Items[23].FieldName := 'whoFrom';
      Columns.Items[24].FieldName := 'sellDate';
      Columns.Items[25].FieldName := 'letterBody';
      Columns.Items[26].FieldName := 'mailName';
      Columns.Items[27].FieldName := 'propertySite';
      Columns.Items[28].FieldName := 'cityZip';
      Columns.Items[29].FieldName := 'dear';
      Columns.Items[30].FieldName := 'inRe';
      Columns.Items[31].FieldName := 'textViolationWords';
      Columns.Items[32].FieldName := 'specificViolationTitle';
      Columns.Items[33].FieldName := 'StreetNumber';
      Columns.Items[34].FieldName := 'StreetName';
      Columns.Items[35].FieldName := 'closing';
      Columns.Items[36].FieldName := 'firstParagraph';
      Columns.Items[37].FieldName := 'legal';
      Columns.Items[38].FieldName := 'acctSectionLine';
      Columns.Items[39].FieldName := 'offsiteAddress';
      Columns.Items[40].FieldName := 'offsiteCityStateZip';
      Columns.Items[41].FieldName := 'closeDate';
  end;
end;


  end.
