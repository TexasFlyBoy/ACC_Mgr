program SSCA_ACDR;

uses
  Forms,
  Main in 'Main.pas' {MainForm},
  About in 'About.pas' {AboutBox},
  SplashScreen in 'SplashScreen.pas' {SplashScn},
  sql_query in 'sql_query.pas',
  ini_file in 'ini_file.pas',
  all_owners in 'all_owners.pas',
  all_letters in 'all_letters.pas',
  genVioLetters in 'genVioLetters.pas',
  AccDb in 'AccDb.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'SSCA Architectural Control Database';
  SplashScn := TSplashScn.Create(Application);
  SplashScn.Show;
  SplashScn.Update;
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TAboutBox, AboutBox);
  while not SplashScn.Completed do
    Application.ProcessMessages;
  SplashScn.Hide;
  SplashScn.Free;
  Application.Run;
  end.
