program SSCA_ACDR;

uses
  Forms,
  Main in 'Main.pas' {MainForm},
  About in 'About.pas' {AboutBox},
  Splash in 'Splash.pas' {FormSplash};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'SSCA Architectural Control Database';
//   Application.CreateForm(TMainForm, MainForm);
//  Application.CreateForm(TAboutBox, AboutBox);
  Application.CreateForm(TFormSplash, FormSplash);
  Application.Run;
  end.
