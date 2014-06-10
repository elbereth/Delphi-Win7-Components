unit FormDemoWin7Controls;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, dwProgressBar, StdCtrls, ExtCtrls, dwTaskbarComponents,
  dwTaskbarThumbnails, ImgList, dwOverlayIcon, dwJumpLists, jpeg;

type
  TfrmWin7Controls = class(TForm)
    tmrProgressBar: TTimer;
    pgcWin7Controls: TPageControl;
    tbsProgress: TTabSheet;
    ProgressBar: TdwProgressBar;
    mmoProgress: TMemo;
    cmbProgressBarState: TComboBox;
    chkMarqueeEnabled: TCheckBox;
    chkShowProgress: TCheckBox;
    chkShowInTaskbar: TCheckBox;
    tbsThumbnails: TTabSheet;
    TaskbarThumbnails: TdwTaskbarThumbnails;
    ilstThumbnailsDefault: TImageList;
    btnThumbnailImagesSwap: TButton;
    ilstThumbnailsAlt: TImageList;
    mmoThumbnails: TMemo;
    edtTaskbarThumbnailHint: TEdit;
    lblThumbnailClicked: TLabel;
    tbsOverlayIcon: TTabSheet;
    ilstOverlayIcons: TImageList;
    cmbOverlay: TComboBoxEx;
    OverlayIcon: TdwOverlayIcon;
    mmoOverlayIcon: TMemo;
    tbsJumpLists: TTabSheet;
    btnRegisterFileHandler: TButton;
    mmoJumpLists: TMemo;
    JumpListsNone: TdwJumpLists;
    JumpListsDefault: TdwJumpLists;
    JumpListsCustom: TdwJumpLists;
    btnUnregisterFileHandler: TButton;
    btnJLNone: TButton;
    lblARJumpLists: TLabel;
    lblMaxJumpListEntryCount: TLabel;
    btnJLDefault: TButton;
    btnJLCustom: TButton;
    btnJLWhole: TButton;
    JumpListsWhole: TdwJumpLists;
    dlgOpenFiles: TOpenDialog;
    btnOpenFile: TButton;
    Image1: TImage;
    tbsAbout: TTabSheet;
    mmoIntro: TMemo;
    lblWin7Needed: TLabel;
    lblCallInfo: TLabel;
    JumpListsDocuments: TdwJumpLists;
    JumpListsTasks: TdwJumpLists;
    btnJLTaskEntries: TButton;
    btnJLCustomDocs: TButton;
    procedure tmrProgressBarTimer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cmbProgressBarStateChange(Sender: TObject);
    procedure chkMarqueeEnabledClick(Sender: TObject);
    procedure chkShowProgressClick(Sender: TObject);
    procedure chkShowInTaskbarClick(Sender: TObject);
    procedure TaskbarThumbnailsThumbnailClick(Sender: TdwTaskbarThumbnailItem);
    procedure btnThumbnailImagesSwapClick(Sender: TObject);
    procedure TaskbarThumbnailEnable(Sender: TObject);
    procedure TaskbarThumbnailBorder(Sender: TObject);
    procedure TaskbarThumbnailDismiss(Sender: TObject);
    procedure TaskbarThumbnailVisible(Sender: TObject);
    procedure TaskbarThumbnailHint(Sender: TObject);
    procedure cmbOverlayChange(Sender: TObject);
    procedure btnRegisterFileHandlerClick(Sender: TObject);
    procedure btnUnregisterFileHandlerClick(Sender: TObject);
    procedure btnJLNoneClick(Sender: TObject);
    procedure btnJLDefaultClick(Sender: TObject);
    procedure btnJLWholeClick(Sender: TObject);
    procedure btnOpenFileClick(Sender: TObject);
    procedure btnJLCustomClick(Sender: TObject);
    procedure btnJLTaskEntriesClick(Sender: TObject);
    procedure btnJLCustomDocsClick(Sender: TObject);
  private
    procedure DoOpenFile(FileName: string);
    procedure AnalyzeParams;
    procedure EnableRegisterButtons;
    procedure DoRegisterFileHandlers(Ext: string; DoRegister: Boolean);
    procedure RegisterRegEntries;
    procedure UnregisterRegEntries;
    function IsFileHandlerRegistered: Boolean;
    procedure FillJumpList(JumpList: TdwJumpLists; ShowTasks: Boolean = True; ShowDocs: Boolean = True);
  public
  end;

var
  frmWin7Controls: TfrmWin7Controls;

implementation

uses
  ShellAPI, ShlObj, ActiveX, ComObj, dwCustomDestinationList, dwObjectArray,
  ComServ, Registry;

{$R *.dfm}
{$R Tasks.res}

const
  SECURITY_NT_AUTHORITY: TSIDIdentifierAuthority = (Value: (0, 0, 0, 0, 0, 5));
  SECURITY_BUILTIN_DOMAIN_RID = $00000020;
  DOMAIN_ALIAS_RID_ADMINS = $00000220;
  SE_GROUP_ENABLED = $00000004;

function IsAdmin: Boolean;
var
  hAccessToken: THandle;
  ptgGroups: PTokenGroups;
  dwInfoBufferSize: DWORD;
  psidAdministrators: PSID;
  x: Integer;
  bSuccess: BOOL;
begin
  Result   := False;
  bSuccess := OpenThreadToken(GetCurrentThread, TOKEN_QUERY, True, hAccessToken);
  if not bSuccess then
    if GetLastError = ERROR_NO_TOKEN then
      bSuccess := OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, hAccessToken);
  if bSuccess then
  begin
    GetTokenInformation(hAccessToken, TokenGroups, nil, 0, dwInfoBufferSize);
    ptgGroups := GetMemory(dwInfoBufferSize); 
    bSuccess := GetTokenInformation(hAccessToken, TokenGroups, ptgGroups, dwInfoBufferSize, dwInfoBufferSize);
    CloseHandle(hAccessToken);
    if bSuccess then
    begin
      AllocateAndInitializeSid(SECURITY_NT_AUTHORITY, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, psidAdministrators);
      {$R-}
      for x := 0 to ptgGroups.GroupCount - 1 do
      begin
        if (SE_GROUP_ENABLED = (ptgGroups.Groups[x].Attributes and SE_GROUP_ENABLED)) and EqualSid(psidAdministrators, ptgGroups.Groups[x].Sid) then
        begin
          Result := True;
          Break;
        end;
      end;
      {$R+}
      FreeSid(psidAdministrators);
    end;
    FreeMem(ptgGroups);
  end;
end;

const
  PROG_ID = 'WiTEC.Samples.Windows7.Controls';

{ TfrmWin7Controls }

procedure TfrmWin7Controls.AnalyzeParams;
begin
  lblCallInfo.Caption := '';
  if ParamStr(1) = '/registerHandlers' then
  begin
    RegisterRegEntries;
  end
  else if ParamStr(1) = '/unregisterHandlers'  then
  begin
    UnregisterRegEntries;
  end else if SameText('/task', Copy(ParamStr(1), 1, 5)) then
  begin
    lblCallInfo.Caption := Format('You started task "%s"', [ParamStr(1)])
  end
  else if SameText('/open', ParamStr(1)) then
  begin
    lblCallInfo.Caption := Format('You opened the file "%s" using param switch "/open"', [ParamStr(2)])
  end
  else if FileExists(ParamStr(1)) then
  begin
    lblCallInfo.Caption := Format('You opened the file "%s" no param switch', [ParamStr(1)])
  end;
  lblCallInfo.Visible := lblCallInfo.Caption <> '';
end;

procedure TfrmWin7Controls.btnJLCustomClick(Sender: TObject);
begin
  JumpListsCustom.Commit;
end;

procedure TfrmWin7Controls.btnJLCustomDocsClick(Sender: TObject);
begin
  JumpListsDocuments.Commit;
end;

procedure TfrmWin7Controls.btnJLDefaultClick(Sender: TObject);
begin
  JumpListsDefault.Commit;
end;

procedure TfrmWin7Controls.btnJLNoneClick(Sender: TObject);
begin
  JumpListsNone.Commit;
end;

procedure TfrmWin7Controls.btnJLTaskEntriesClick(Sender: TObject);
begin
  JumpListsTasks.Commit;
end;

procedure TfrmWin7Controls.btnJLWholeClick(Sender: TObject);
begin
  JumpListsWhole.Commit;
end;

procedure TfrmWin7Controls.btnOpenFileClick(Sender: TObject);
begin
  if dlgOpenFiles.Execute then
    DoOpenFile(dlgOpenFiles.FileName);
end;

procedure TfrmWin7Controls.btnRegisterFileHandlerClick(Sender: TObject);
begin
  if not IsAdmin then
  begin
    if CheckWin32Version(6, 0) then
    begin
      if ShellExecute(Handle, 'runas', PChar(Application.Exename), '/registerHandlers', nil, SW_SHOWNORMAL) > 32 then
        Application.Terminate;
    end;
    Exit;
  end;

  RegisterRegEntries;
  EnableRegisterButtons;
end;

procedure TfrmWin7Controls.btnThumbnailImagesSwapClick(Sender: TObject);
begin
  if TaskbarThumbnails.Images = ilstThumbnailsAlt then
    TaskbarThumbnails.Images := ilstThumbnailsDefault
  else
    TaskbarThumbnails.Images := ilstThumbnailsAlt;
end;

procedure TfrmWin7Controls.btnUnregisterFileHandlerClick(Sender: TObject);
begin
  if not IsAdmin then
  begin
    if CheckWin32Version(6, 0) then
    begin
      if ShellExecute(Handle, 'runas', PChar(Application.Exename), '/unregisterHandlers', nil, SW_SHOWNORMAL) > 32 then
        Application.Terminate;
    end;
    Exit;
  end;
  
  UnregisterRegEntries;
  EnableRegisterButtons;
end;

function SetCurrentProcessExplicitAppUserModelID(AppID: LPCWSTR): HResult; stdcall; external 'Shell32.dll';

procedure TfrmWin7Controls.chkMarqueeEnabledClick(Sender: TObject);
begin
  ProgressBar.MarqueeEnabled := chkMarqueeEnabled.Checked;
end;

procedure TfrmWin7Controls.chkShowInTaskbarClick(Sender: TObject);
begin
  ProgressBar.ShowInTaskbar := chkShowInTaskbar.Checked;
end;

procedure TfrmWin7Controls.chkShowProgressClick(Sender: TObject);
begin
  tmrProgressBar.Enabled := chkShowProgress.Checked;
end;

procedure TfrmWin7Controls.cmbOverlayChange(Sender: TObject);
begin
  OverlayIcon.ImageIndex := cmbOverlay.ItemsEx[cmbOverlay.ItemIndex].ImageIndex;
  OverlayIcon.Hint := cmbOverlay.ItemsEx[cmbOverlay.ItemIndex].Caption;
end;

procedure TfrmWin7Controls.cmbProgressBarStateChange(Sender: TObject);
begin
  ProgressBar.ProgressBarState := TdwProgressBarState(cmbProgressBarState.ItemIndex);
  chkShowProgress.Checked := False;
end;

procedure TfrmWin7Controls.DoOpenFile(FileName: string);
begin
  ShowMessage(Format('File "%s" opened, okay not ;-)', [Filename]));
end;

procedure TfrmWin7Controls.DoRegisterFileHandlers(Ext: string; DoRegister: Boolean);
var
  Reg: TRegistry;
  Key: string;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    Key := Ext + '\OpenWithProgids';
    if Reg.OpenKey(Key, True) then
    begin
      if DoRegister then
        Reg.WriteString(PROG_ID, '')
      else
        Reg.DeleteValue(PROG_ID);
    end;
  finally
    Reg.Free;
  end;
end;

procedure TfrmWin7Controls.EnableRegisterButtons;
const
  BCM_FIRST = $1600;
  BCM_SETSHIELD = BCM_FIRST + $000C;
begin
  btnRegisterFileHandler.Enabled := True;
  btnUnregisterFileHandler.Enabled := True;

  if CheckWin32Version(6, 0) then
  begin
    if IsAdmin then
    begin
      lblARJumpLists.Visible := False;
      SendMessage(btnRegisterFileHandler.Handle, BCM_SETSHIELD, 0, 0);
      SendMessage(btnUnregisterFileHandler.Handle, BCM_SETSHIELD, 0, 0);
    end
    else
    begin
      lblARJumpLists.Visible := True;
      SendMessage(btnRegisterFileHandler.Handle, BCM_SETSHIELD, 0, 1);
      SendMessage(btnUnregisterFileHandler.Handle, BCM_SETSHIELD, 0, 1);
    end;
  end
  else
  begin
    btnRegisterFileHandler.Enabled := IsAdmin;
    btnUnregisterFileHandler.Enabled := IsAdmin;
    lblARJumpLists.Visible := not IsAdmin;
  end;

  btnRegisterFileHandler.Enabled := btnRegisterFileHandler.Enabled and (not IsFileHandlerRegistered);
  btnUnregisterFileHandler.Enabled := btnUnregisterFileHandler.Enabled and IsFileHandlerRegistered;

  btnJLCustom.Enabled := IsFileHandlerRegistered;
  btnJLWhole.Enabled := IsFileHandlerRegistered;
end;

procedure TfrmWin7Controls.FillJumpList(JumpList: TdwJumpLists; ShowTasks, ShowDocs: Boolean);
var
  Category: TdwLinkCategoryItem;
begin
  if ShowTasks then
  begin
    JumpList.Tasks.AddShellLink('Run Task 1', '/task1', Application.ExeName, 1);
    JumpList.Tasks.AddShellLink('Run Task 2', '/task2', Application.ExeName, 2);
    JumpList.Tasks.AddShellLink('Run Task 3', '/task3', Application.ExeName, 3);
  end;

  if ShowDocs then
  begin
    Category := JumpList.Categories.Add;
    Category.Title := 'Using Shell Items';
    Category.Items.AddShellItem(ExtractFilePath(Application.ExeName) + 'sample1.w7c');
    Category.Items.AddShellItem(ExtractFilePath(Application.ExeName) + 'sample2.w7c');

    Category := JumpList.Categories.Add;
    Category.Title := 'Using Shell Links';
    Category.Items.AddShellLink('Open just a document', '/open "' + ExtractFilePath(Application.ExeName) + 'sample3.w7c"', Application.ExeName, 4);
  end;
end;

procedure TfrmWin7Controls.FormCreate(Sender: TObject);
begin
  pgcWin7Controls.ActivePageIndex := 0;
  mmoIntro.WordWrap := True;
  lblWin7Needed.Visible := not CheckWin32Version(6, 1);
  mmoProgress.WordWrap := True;
  mmoThumbnails.WordWrap := True;
  mmoOverlayIcon.WordWrap := True;
  cmbOverlay.ItemIndex := 0;
  mmoJumpLists.WordWrap := True;
  lblMaxJumpListEntryCount.Caption := Format('Maximum Jump List Entry Count: %d', [JumpListsNone.GetMaxJumpListEntryCount]);

  AnalyzeParams;

  EnableRegisterButtons;

  FillJumpList(JumpListsCustom);
  FillJumpList(JumpListsWhole);
  FillJumpList(JumpListsTasks, True, False);
  FillJumpList(JumpListsDocuments, False);
end;

function TfrmWin7Controls.IsFileHandlerRegistered: Boolean;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    Result := Reg.OpenKey(PROG_ID, False);
  finally
    Reg.Free;
  end;
end;

procedure TfrmWin7Controls.RegisterRegEntries;
var
  Reg: TRegistry;
begin
  if not IsAdmin then
    Exit;

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    Reg.OpenKey(PROG_ID, True);
    Reg.WriteString('', 'Windows 7 Controls Demo');
    Reg.OpenKey('Shell\Open\Command', True);
    Reg.WriteString('', Application.ExeName + ' /open %1');
  finally
    Reg.Free;
  end;

  DoRegisterFileHandlers('.txt', True);
  DoRegisterFileHandlers('.w7c', True);
end;

procedure TfrmWin7Controls.TaskbarThumbnailBorder(Sender: TObject);
begin
  if Sender = nil then
    Exit;
  if not (Sender is TCheckBox) then
    Exit;

  with Sender as TCheckBox do
  begin
    TaskbarThumbnails.Thumbnails[Tag].ShowBorder := Checked;
  end;
end;

procedure TfrmWin7Controls.TaskbarThumbnailDismiss(Sender: TObject);
begin
  if Sender = nil then
    Exit;
  if not (Sender is TCheckBox) then
    Exit;

  with Sender as TCheckBox do
  begin
    TaskbarThumbnails.Thumbnails[Tag].DismissOnClick := Checked;
  end;
end;

procedure TfrmWin7Controls.TaskbarThumbnailEnable(Sender: TObject);
begin
  if Sender = nil then
    Exit;
  if not (Sender is TCheckBox) then
    Exit;

  with Sender as TCheckBox do
  begin
    TaskbarThumbnails.Thumbnails[Tag].Enabled := Checked;
  end;
end;

procedure TfrmWin7Controls.TaskbarThumbnailHint(Sender: TObject);
begin
  if Sender = nil then
    Exit;
  if not (Sender is TEdit) then
    Exit;

  with Sender as TEdit do
  begin
    TaskbarThumbnails.Thumbnails[Tag].Hint := Text;
  end;
end;

procedure TfrmWin7Controls.TaskbarThumbnailsThumbnailClick(Sender: TdwTaskbarThumbnailItem);
begin
  lblThumbnailClicked.Caption := Format('Button: %d, %s', [Sender.Index, Sender.Hint]);
end;

procedure TfrmWin7Controls.TaskbarThumbnailVisible(Sender: TObject);
begin
  if Sender = nil then
    Exit;
  if not (Sender is TCheckBox) then
    Exit;

  with Sender as TCheckBox do
  begin
    TaskbarThumbnails.Thumbnails[Tag].Visible := Checked;
  end;
end;

procedure TfrmWin7Controls.tmrProgressBarTimer(Sender: TObject);
begin
  if ProgressBar.Position = ProgressBar.Max then
    ProgressBar.Position := 0
  else
    ProgressBar.StepIt;

  if ProgressBar.Position = ProgressBar.Max then
    tmrProgressBar.Interval := 2500
  else
    tmrProgressBar.Interval := 20;
end;

procedure TfrmWin7Controls.UnregisterRegEntries;
var
  Reg: TRegistry;
begin
  if not IsAdmin then
    Exit;

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    Reg.DeleteKey(PROG_ID);
  finally
    Reg.Free;
  end;

  DoRegisterFileHandlers('.txt', False);
  DoRegisterFileHandlers('.w7c', False);
end;

initialization
//  Randomize;
//  SetCurrentProcessExplicitAppUserModelID(PWideChar(WideString(IntToStr(Random(5)))));
//  TComObjectFactory.Create(ComServer, TMyItemsArray, IID_IObjectArray, 'MyItems', 'Desc', ciInternal, tmSingle);

end.
