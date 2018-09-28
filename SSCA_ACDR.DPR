program SSCA_ACDR;

uses
  Forms,
  Main in 'Main.pas' {MainForm},
  About in 'About.pas' {AboutBox};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'SSCA Architectural Control Database';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
  end.
