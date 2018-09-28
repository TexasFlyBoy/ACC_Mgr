unit CurrentOwnersFrame;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, 
  Dialogs, DB, ADODB, StdCtrls, Grids, DBGrids, Buttons, ComCtrls;

type
  TFrame1 = class(TFrame)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    frm1DbGrid: TDBGrid;
    eOwnersOwner: TEdit;
    frame1AdoTableCurrentOwners: TADOTable;
    frame1AdoConnection: TADOConnection;
    dsCurrentOwners: TDataSource;
    sbAcctSort: TSpeedButton;
    sbOwnerSort: TSpeedButton;
    sbStrNumSort: TSpeedButton;
    sbStrNameSort: TSpeedButton;
    sbCurrentOwnersFrame: TStatusBar;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    procedure EditColorChangeOnEnter(Sender: TObject);
    procedure EditColorChangeOnExit(Sender: TObject);
    procedure sbAcctSortClick(Sender: TObject);
    procedure sbOwnerSortClick(Sender: TObject);
    procedure sbStrNumSortClick(Sender: TObject);
    procedure sbStrNameSortClick(Sender: TObject);
    procedure eOwnersOwnerChange(Sender: TObject);
    procedure frm1DbGridDblClick(Sender: TObject);
    procedure eOwnersHouseAcctChange(Sender: TObject);
    procedure eOwnersStreetSearchChange(Sender: TObject);
    procedure eOwnersAddressSearchChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.dfm}

procedure TFrame1.EditColorChangeOnEnter(Sender: TObject);
begin
  TEdit(Sender).Color := clMoneyGreen;
end;

procedure TFrame1.EditColorChangeOnExit(Sender: TObject);
begin
  TEdit(Sender).Color := clWindow;
end;

procedure TFrame1.sbAcctSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  frame1AdoTableCurrentOwners.Filtered := False;
  frame1AdoTableCurrentOwners.IndexFieldNames := 'houseAcct';
  frame1AdoTableCurrentOwners.Filtered := True;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TFrame1.sbOwnerSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  frame1AdoTableCurrentOwners.Filtered := False;
  frame1AdoTableCurrentOwners.IndexFieldNames := 'Owner';
  frame1AdoTableCurrentOwners.Filtered := True;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TFrame1.sbStrNumSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  frame1AdoTableCurrentOwners.Filtered := False;
  frame1AdoTableCurrentOwners.IndexFieldNames := 'streetNumber';
  frame1AdoTableCurrentOwners.Filtered := True;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TFrame1.sbStrNameSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  frame1AdoTableCurrentOwners.Filtered := False;
  frame1AdoTableCurrentOwners.IndexFieldNames := 'streetName';
  frame1AdoTableCurrentOwners.Filtered := True;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TFrame1.eOwnersOwnerChange(Sender: TObject);
var
  numOfRecords: integer;
begin
  if (length(eOwnersOwner.Text) = 0) then begin
      frame1AdoTableCurrentOwners.Filter := '';
      frame1AdoTableCurrentOwners.Filtered := True;
      end
  else begin
      frame1AdoTableCurrentOwners.Filtered := False;
      frame1AdoTableCurrentOwners.Filter := 'Owner like ' + '''' + eOwnersOwner.text + '%' + '''' ;
      frame1AdoTableCurrentOwners.Filtered := true;
  end;
  sbCurrentOwnersFrame.Panels[0].Text := frame1AdoTableCurrentOwners.Filter;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TFrame1.frm1DbGridDblClick(Sender: TObject);
var
  filterHouseAcct: integer;
begin
  ShowMessage('2222');
  filterHouseAcct := frame1AdoTableCurrentOwners.FieldValues['houseAcct'];
  frame1AdoTableCurrentOwners.Filtered := false;
  frame1AdoTableCurrentOwners.Filter := 'houseAcct = ' + IntToStr(filterHouseAcct) ;
  frame1AdoTableCurrentOwners.Filtered := true;
  ShowMessage('3333');
end;

procedure TFrame1.eOwnersHouseAcctChange(Sender: TObject);
var
  numOfRecords: integer;
begin
  if (length(Edit1.Text) = 0) then begin
      frame1AdoTableCurrentOwners.Filter := '';
      frame1AdoTableCurrentOwners.Filtered := True;
      end
  else begin
      frame1AdoTableCurrentOwners.Filtered := False;
      frame1AdoTableCurrentOwners.Filter := 'houseAcct = ' + Edit1.text ;
      frame1AdoTableCurrentOwners.Filtered := true;
  end;
  sbCurrentOwnersFrame.Panels[0].Text := frame1AdoTableCurrentOwners.Filter;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TFrame1.eOwnersStreetSearchChange(Sender: TObject);
var
  numOfRecords: integer;
begin
  if (length(edit2.Text) = 0) then begin
      frame1AdoTableCurrentOwners.Filter := '';
      frame1AdoTableCurrentOwners.Filtered := True;
      end
  else begin
      frame1AdoTableCurrentOwners.Filtered := False;
      frame1AdoTableCurrentOwners.Filter := 'streetName like ' + '''' + edit2.text + '%' + '''' ;
      frame1AdoTableCurrentOwners.Filtered := true;
  end;
  sbCurrentOwnersFrame.Panels[0].Text := frame1AdoTableCurrentOwners.Filter;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TFrame1.eOwnersAddressSearchChange(Sender: TObject);
var
  numOfRecords: integer;
begin
  if (length(edit3.Text) = 0) then begin
      frame1AdoTableCurrentOwners.Filter := '';
      frame1AdoTableCurrentOwners.Filtered := True;
      end
  else begin
      frame1AdoTableCurrentOwners.Filtered := False;
      frame1AdoTableCurrentOwners.Filter := 'address like ' + '''' + edit3.text + '%' + '''' ;
      frame1AdoTableCurrentOwners.Filtered := true;
  end;
  sbCurrentOwnersFrame.Panels[0].Text := frame1AdoTableCurrentOwners.Filter;
  numOfRecords := frame1AdoTableCurrentOwners.RecordCount;
  sbCurrentOwnersFrame.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;
end.
