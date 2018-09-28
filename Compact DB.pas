{
After closing the TADOConnection and ALL DataSets associated with it, you need to make sure the db is unlocked. Remember that other users might be connected to the db and in that case you cannot compact it.

Before actually compressing the db you have to give the jet engine a bit of time to actually close the connection, flush, and unlock the db. Then test if the db is locked (try to open for exclusive use).

Here is the method I use, which always worked for me:
}


uses ComObj;

procedure JroRefreshCache(ADOConnection: TADOConnection);
var
  JetEngine: OleVariant;
begin
  if not ADOConnection.Connected then Exit;
  JetEngine := CreateOleObject('jro.JetEngine');
  JetEngine.RefreshCache(ADOConnection.ConnectionObject);
end;

procedure JroCompactDatabase(const Source, Destination: string);
var
  JetEngine: OleVariant;
begin
  JetEngine := CreateOleObject('jro.JetEngine');
  JetEngine.CompactDatabase(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + Source,
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + Destination + ';Jet OLEDB:Engine Type=5');
end;

procedure CompactDatabase(const MdbFileName: string;
  ADOConnection: TADOConnection=nil;
  const ReopenConnection: Boolean=True);
var
  LdbFileName, TempFileName: string;
  FailCount: Integer;
  FileHandle: Integer;
begin
  TempFileName := ChangeFileExt(MdbFileName, '.temp.mdb');
  if Assigned(ADOConnection) then
  begin
    // force the database engine to write data to disk, releasing locks on memory
    JroRefreshCache(ADOConnection);
    // close the connection - this will also close all associated datasets
    ADOConnection.Close;
  end;
  // ADOConnection.Close SHOULD delete the ldb
  // force delete of ldb lock file just in case if we don't have an active ADOConnection
  LdbFileName := ChangeFileExt(MdbFileName, '.ldb');
  if FileExists(LdbFileName) then
    DeleteFile(LdbFileName); // could fail because data is still locked - we ignore this
  // delete temp file if any
  if FileExists(TempFileName) then
    if not DeleteFile(TempFileName) then
       RaiseLastOSError;
  // try to open for exclusive use
  FailCount := 0;
  repeat
    FileHandle := FileOpen(MdbFileName, fmShareExclusive);
    try
      if FileHandle = -1 then // error
      begin 
        Inc(FailCount);
        Sleep(100); // give the database engine time to close completely and unlock
      end
      else
      begin
        FailCount := 0;
        Break; // success
      end;
    finally
      FileClose(FileHandle);
    end;
  until FailCount = 10; // maximum 1 second of attempts      
  if FailCount <> 0 then // file is probably locked by another user/process        
    raise Exception.Create(Format('Error opening %s for exclusive use.', [MdbFileName]));
  // compact the db
  JroCompactDatabase(MdbFileName, TempFileName);
  // copy temp file to original mdb and delete temp file on success
  if Windows.CopyFile(PChar(TempFileName), PChar(MdbFileName), False) then
    DeleteFile(TempFileName)
  else
    RaiseLastOSError;
  // reopen ADOConnection
  if Assigned(ADOConnection) and ReopenConnection then
    ADOConnection.Open;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  CompactDatabase('F:\Projects\DB\mydb.mdb', ADOConnection1, True);
  // reopen DataSets
  ADODataSet1.Open;
end;

{
Make sure that your TADOConnection is NOT set to Connected in the IDE (Design mode).
Because if it does, there is another active connection to the db.

edited Nov 15 '13 at 1:43
answered Nov 14 '13 at 23:28

kobik



----------------------------------------------------------

I tried your code but I still get if FailCount <> 0 then // file is probably locked by another user/process raise Exception.Create(Format('Error opening %s for exclusive use.', [MdbFileName])); Exception is raised in CompactDatabase. I also tried to increase FailCount from 10 to 100. My database is on a non-networked pc with only this app accessing the database. 

– Bill Nov 15 '13 at 0:34


My previous comment noted an exception while the app was running in the XE4 IDE. I just tried the app by executing the EXE and the database was compacted with no exception? 

– Bill Nov 15 '13 at 1:34

----------------------------------------------------------

My guess is that your ADOConnection1 is set to Connected in the IDE? That means that another user is connected already. 

– kobik Nov 15 '13 at 1:35 

}