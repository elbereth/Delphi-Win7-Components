unit uWin7Registration;

interface

procedure Register;

{$R ToolPaletteIcons.dcr}

implementation

uses
  Classes,
  dwProgressBar, dwTaskbarThumbnails, dwOverlayIcon, dwJumpLists;

procedure Register;
begin
  RegisterComponents('Windows 7 Support', [TdwProgressBar, TdwTaskbarThumbnails, TdwOverlayIcon, TdwJumpLists]);
end;

end.
