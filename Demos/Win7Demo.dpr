program Win7Demo;

uses
  Forms,
  FormDemoWin7Controls in 'FormDemoWin7Controls.pas' {frmWin7Controls},
  dwProgressBar in '..\dwProgressBar\dwProgressBar.pas',
  dwTaskbarComponents in '..\dwCommon\dwTaskbarComponents.pas',
  dwTaskbarList in '..\dwCommon\dwTaskbarList.pas',
  dwTaskbarThumbnails in '..\dwTaskbarThumbnails\dwTaskbarThumbnails.pas',
  dwOverlayIcon in '..\dwOverlayIcon\dwOverlayIcon.pas',
  dwCustomDestinationList in '..\dwCommon\dwCustomDestinationList.pas',
  dwObjectArray in '..\dwCommon\dwObjectArray.pas',
  dwJumpLists in '..\dwJumpLists\dwJumpLists.pas';

{$R *.res}

{$INCLUDE '..\Packages\DelphiVersions.inc'}

begin
  Application.Initialize;
  {$IFDEF Delphi2007_Up}
    Application.MainFormOnTaskbar := True;
  {$ENDIF}
  Application.Title := 'Windows 7 Control Support Demo';
  Application.CreateForm(TfrmWin7Controls, frmWin7Controls);
  Application.Run;
end.
