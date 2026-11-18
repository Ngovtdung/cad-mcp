; Setup cho "Claude AI cho AutoCAD". Biên dịch qua installer\pack_setup.py (/DAppVer /DStage /DOutDir). Spec: docs\spec-installer.md §4.
#ifndef AppVer
  #error Thieu /DAppVer
#endif
#ifndef Stage
  #error Thieu /DStage
#endif
#ifndef OutDir
  #define OutDir "."
#endif
#define AppGuid "{E16EAF7B-C860-4636-BC5D-8CA9A1004C51}"

[Setup]
AppId={code:AppIdValue}
UsePreviousLanguage=no
AppName=Claude AI cho AutoCAD
AppVersion={#AppVer}
AppVerName=Claude AI cho AutoCAD {#AppVer}
AppPublisher=cad-mcp
AppCopyright=Copyright (c) 2026 Gaus
VersionInfoVersion={#AppVer}
DefaultDirName={localappdata}\cad-mcp
DisableDirPage=yes
UsePreviousAppDir=no
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
CloseApplications=no
RestartApplications=no
OutputDir={#OutDir}
OutputBaseFilename=cad-mcp-setup-{#AppVer}
SetupIconFile={#Stage}\app\app_icon.ico
UninstallDisplayIcon={app}\app_icon.ico
UninstallDisplayName={code:DisplayName}
InfoBeforeFile={#Stage}\setup\truoc-khi-cai.txt
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
ShowLanguageDialog=no

[Languages]
Name: "vi"; MessagesFile: "Vietnamese.isl"

[Tasks]
Name: "claudecli"; Description: "Cài Claude Code (tải claude.exe từ downloads.claude.ai, chỉ chạy nếu chữ ký là Anthropic, PBC)"; Check: NeedClaudeCli

[Files]
Source: "{#Stage}\app\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#Stage}\bundle\*"; DestDir: "{code:BundleDir}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#Stage}\setup\claude-cli.ps1"; Flags: dontcopy

[InstallDelete]
; lối tắt cũ của 3.0.0 (tên đã đổi)
Type: files; Name: "{autoprograms}\Claude AI cho AutoCAD\Hướng dẫn sử dụng.lnk"; Check: not IsTest
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\vendor"
Type: files; Name: "{code:BundleDir}\Contents\ClaudePalette.dll"

[UninstallDelete]
; KHÔNG filesandordirs trên {app}. goi_y.txt được GIỮ (Q3) -> {app} còn lại nếu có goi_y.txt
Type: filesandordirs; Name: "{app}\python"
Type: filesandordirs; Name: "{app}\vendor"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: files; Name: "{app}\mcp.json"
Type: filesandordirs; Name: "{app}\logs"
Type: dirifempty; Name: "{app}"
Type: dirifempty; Name: "{autoprograms}\Claude AI cho AutoCAD"

[Registry]
; không cờ uninsdelete: CurUninstallStepChanged quyết theo Home
Root: HKCU; Subkey: "Software\{code:RegKeyName}"; ValueType: string; ValueName: "Home"; ValueData: "{app}"

[Icons]
Name: "{autoprograms}\Claude AI cho AutoCAD\Gỡ cài đặt Claude AI cho AutoCAD"; Filename: "{uninstallexe}"; Check: not IsTest
Name: "{autoprograms}\Claude AI cho AutoCAD\Hướng dẫn sử dụng Claude AI cho AutoCAD"; Filename: "{app}\docs\HDSD.html"; Check: not IsTest

[Run]
Filename: "{code:VerifiedClaudeExe}"; Parameters: "auth login"; Description: "Đăng nhập Claude ngay (mở cửa sổ đăng nhập)"; Flags: postinstall nowait skipifsilent; Check: OfferLogin
Filename: "{app}\docs\HDSD.html"; Description: "Mở Hướng dẫn sử dụng"; Flags: postinstall shellexec nowait skipifsilent

[Code]
var
  VerifiedClaude: String;   { claude.exe đã kiểm chữ ký Anthropic ở ssPostInstall; '' = không có }
  LoggedIn: Boolean;        { kết quả -Mode status (mã 0) chạy 1 lần ở cuối ssPostInstall; timeout/lỗi = False }
  CliExtracted: Boolean;
  RiskPage: TWizardPage;
  RiskCheck: TNewCheckBox;

function IsDriveAbs(S: String): Boolean;
begin
  Result := (Length(S) >= 3) and (S[2] = ':') and (S[3] = '\') and
    (((S[1] >= 'A') and (S[1] <= 'Z')) or ((S[1] >= 'a') and (S[1] <= 'z')));
end;

function Norm(S: String): String;
begin
  Result := RemoveBackslash(ExpandFileName(S));
end;

function IsTest: Boolean;
begin
  Result := ExpandConstant('{param:TESTROOT|}') <> '';
end;

function TestRoot: String;
begin
  Result := Norm(ExpandConstant('{param:TESTROOT|}'));
end;

function AppIdValue(Param: String): String;
begin
  Result := '{#AppGuid}';
  if IsTest then Result := Result + '-TEST';
end;

function DisplayName(Param: String): String;
begin
  Result := 'Claude AI cho AutoCAD';
  if IsTest then Result := Result + ' (TEST)';
end;

function RegKeyName(Param: String): String;
begin
  if IsTest then Result := 'ClaudeCad-Test' else Result := 'ClaudeCad';
end;

function BundleDir(Param: String): String;
begin
  if IsTest then
    Result := TestRoot + '\ApplicationPlugins\ClaudeCAD.bundle'
  else
    Result := ExpandConstant('{userappdata}\Autodesk\ApplicationPlugins\ClaudeCAD.bundle');
end;

{ SEC-2: giống Core.FindClaude — ~\.local\bin rồi từng mục PATH (bỏ ", trim), chỉ "X:\..." }
function ClaudeExe: String;
var
  P, D: String;
  I: Integer;
begin
  Result := ExpandConstant('{%USERPROFILE}\.local\bin\claude.exe');
  if IsDriveAbs(Result) and FileExists(Result) then Exit;
  P := GetEnv('PATH') + ';';
  while P <> '' do
  begin
    I := Pos(';', P);
    D := Copy(P, 1, I - 1);
    Delete(P, 1, I);
    StringChangeEx(D, '"', '', True);
    D := Trim(D);
    if IsDriveAbs(D) then
    begin
      Result := AddBackslash(D) + 'claude.exe';
      if FileExists(Result) then Exit;
    end;
  end;
  Result := '';
end;

function VerifiedClaudeExe(Param: String): String;
begin
  Result := VerifiedClaude;
end;

function NeedClaudeCli: Boolean;
begin
  Result := (not IsTest) and (ClaudeExe = '');
end;

function ListHas(L, Item: String): Boolean;
var
  I: Integer;
  T: String;
begin
  Result := False;
  L := Lowercase(L) + ',';
  while L <> '' do
  begin
    I := Pos(',', L);
    T := Trim(Copy(L, 1, I - 1));
    Delete(L, 1, I);
    if T = Item then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

{ SEC-1: cài im lặng chỉ tải Claude khi /TASKS hoặc /MERGETASKS ghi rõ claudecli }
function SilentWantsClaude: Boolean;
begin
  Result := ListHas(ExpandConstant('{param:TASKS|}'), 'claudecli') or ListHas(ExpandConstant('{param:MERGETASKS|}'), 'claudecli');
end;

function PsExe: String;
begin
  Result := ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe');
end;

function RunCliScript(Params: String; var Rc: Integer): Boolean;
begin
  if not CliExtracted then
  begin
    ExtractTemporaryFile('claude-cli.ps1');
    CliExtracted := True;
  end;
  Result := ExecAndLogOutput(PsExe, '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' +
    ExpandConstant('{tmp}\claude-cli.ps1') + '" ' + Params, '', SW_HIDE, ewWaitUntilTerminated, Rc, nil);
end;

function HasBadQuote(S: String): Boolean;
var
  I: Integer;
begin
  Result := Pos('"', S) > 0;
  for I := 1 to Length(S) do
    if (Ord(S[I]) >= $2018) and (Ord(S[I]) <= $201B) then Result := True;
end;

{ SEC-2: fail closed với ngoặc kép kiểu chữ / dấu "; đường dẫn đi qua argv (-File), không thành mã PowerShell }
function ClaudeSigned(Exe: String): Boolean;
var
  Rc: Integer;
begin
  Result := False;
  if HasBadQuote(Exe) then
  begin
    Log('Tu choi duong dan claude.exe co dau ngoac');
    Exit;
  end;
  Result := RunCliScript('-Mode verify -Path "' + Exe + '"', Rc) and (Rc = 0);
end;

function AcadRunning: Boolean;
var
  Loc, Svc, Res: Variant;
  Py: String;
begin
  Result := False;
  try
    { M3: python.exe nằm dưới thư mục python của bản cài cũng giữ tệp. WQL: \ gấp đôi, ' -> \'; LIKE: [ -> [[] và _ -> [_] (% không có trong đường dẫn Windows hợp lệ ở đây). Thứ tự: [ trước, _ sau, vì [_] chứa [ }
    Py := Norm(ExpandConstant('{app}')) + '\python\';
    StringChangeEx(Py, '\', '\\', True);
    StringChangeEx(Py, '''', '\''', True);
    StringChangeEx(Py, '[', '[[]', True);
    StringChangeEx(Py, '_', '[_]', True);
    Loc := CreateOleObject('WbemScripting.SWbemLocator');
    Svc := Loc.ConnectServer('.', 'root\CIMV2');
    Res := Svc.ExecQuery('SELECT ProcessId FROM Win32_Process WHERE Name = ''acad.exe'' OR (Name = ''python.exe'' AND ExecutablePath LIKE ''' + Py + '%'')');
    Result := Res.Count > 0;
  except
    Log('Khong kiem tra duoc AutoCAD dang chay: ' + GetExceptionMessage);
    if IsUninstaller then
      if UninstallSilent then Result := True;  { gỡ im lặng mà không kiểm được: coi là đang chạy (từ chối), không gỡ mù }
  end;
end;

function AcadList: String;
var
  Rels, Prods: TArrayOfString;
  I, J: Integer;
  K, Name, Loc: String;
begin
  Result := '';
  if not RegGetSubkeyNames(HKLM64, 'SOFTWARE\Autodesk\AutoCAD', Rels) then Exit;
  for I := 0 to GetArrayLength(Rels) - 1 do
  begin
    if (Copy(Rels[I], 1, 1) <> 'R') or (Pos('.', Rels[I]) < 3) then Continue;
    if StrToIntDef(Copy(Rels[I], 2, Pos('.', Rels[I]) - 2), 0) < 22 then Continue;
    K := 'SOFTWARE\Autodesk\AutoCAD\' + Rels[I];
    if not RegGetSubkeyNames(HKLM64, K, Prods) then Continue;
    for J := 0 to GetArrayLength(Prods) - 1 do
      if (Pos(':', Prods[J]) > 0) and RegQueryStringValue(HKLM64, K + '\' + Prods[J], 'AcadLocation', Loc)
         and FileExists(AddBackslash(Loc) + 'acad.exe') then
      begin
        if not RegQueryStringValue(HKLM64, K + '\' + Prods[J], 'ProductName', Name) then Name := Rels[I];
        Result := Result + '  - ' + Name + #13#10;
      end;
  end;
end;

function RiskText: String;
begin
  Result :=
    'Palette cho phép Claude TỰ ĐỘNG dùng mọi công cụ bản vẽ, gồm chạy lệnh AutoCAD và chạy mã Python trên máy này với quyền của bạn, không hỏi lại.' + #13#10#13#10 +
    'Chữ trong bản vẽ hoặc tệp đính kèm (ví dụ DWG nhận từ người khác) có thể chứa chỉ dẫn độc khiến Claude chạy mã: đọc, sửa, xoá tệp, ' +
    'đọc thông tin đăng nhập Claude (%USERPROFILE%\.claude\.credentials.json), gửi dữ liệu ra ngoài. Ảnh hưởng kéo dài tới khi bấm "Mới".' + #13#10#13#10 +
    'Nội dung bản vẽ, yêu cầu và tệp đính kèm được gửi tới Anthropic.' + #13#10#13#10 +
    'Chỉ dùng với bản vẽ và tệp tin cậy; lưu bản vẽ trước khi dùng.';
end;

procedure InitializeWizard;
var
  Memo: TNewMemo;
begin
  RiskPage := CreateCustomPage(wpInfoBefore, 'Rủi ro bảo mật — đọc kỹ trước khi cài', 'Palette cho Claude tự chạy lệnh và mã trên máy bạn.');
  Memo := TNewMemo.Create(RiskPage);
  Memo.Parent := RiskPage.Surface;
  Memo.Left := 0;
  Memo.Top := 0;
  Memo.Width := RiskPage.SurfaceWidth;
  Memo.Height := RiskPage.SurfaceHeight - ScaleY(36);
  Memo.ReadOnly := True;
  Memo.WordWrap := True;
  Memo.ScrollBars := ssVertical;
  Memo.Text := RiskText;
  RiskCheck := TNewCheckBox.Create(RiskPage);
  RiskCheck.Parent := RiskPage.Surface;
  RiskCheck.Left := 0;
  RiskCheck.Top := Memo.Top + Memo.Height + ScaleY(10);
  RiskCheck.Width := RiskPage.SurfaceWidth;
  RiskCheck.Height := ScaleY(20);
  RiskCheck.Caption := 'Tôi hiểu rủi ro trên và chỉ dùng palette với bản vẽ, tệp tin cậy.';
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  { im lặng: Inno vẫn gọi hàm này cho từng trang -> MsgBox treo; bắt buộc /ACCEPTRISK=yes đã kiểm ở InitializeSetup }
  if WizardSilent then Exit;
  if (RiskPage <> nil) and (CurPageID = RiskPage.ID) and not RiskCheck.Checked then
  begin
    MsgBox('Hãy đánh dấu ô "Tôi hiểu rủi ro…" để tiếp tục.', mbError, MB_OK);
    Result := False;
  end;
end;

function InitializeSetup: Boolean;
var
  L: String;
begin
  Result := True;
  if WizardSilent and (CompareText(ExpandConstant('{param:ACCEPTRISK|}'), 'yes') <> 0) then
  begin
    Log('Cai im lang can /ACCEPTRISK=yes (xem Huong dan su dung, muc 6 va 9)');
    Result := False;
    Exit;
  end;
  if IsTest then Exit;
  L := AcadList;
  Log('AutoCAD 2018+ tim thay:'#13#10 + L);
  if (L = '') and not WizardSilent then
    Result := MsgBox('Không tìm thấy AutoCAD 2018 trở lên (bản đầy đủ, 64-bit) trên máy này. AutoCAD LT không được hỗ trợ.'#13#10#13#10 +
      'Vẫn tiếp tục cài?', mbConfirmation, MB_YESNO) = IDYES;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  App, Old, R: String;
begin
  Result := '';
  App := Norm(ExpandConstant('{app}'));
  if FileExists(App + '\ClaudePalette.cs') then
  begin
    Result := 'Không cài đè lên thư mục mã nguồn: ' + App;
    Exit;
  end;
  if IsTest then
  begin
    R := TestRoot;
    if not IsDriveAbs(ExpandConstant('{param:TESTROOT|}')) or not DirExists(R) or not FileExists(R + '\.cad-mcp-testroot') then
      Result := 'TESTROOT không hợp lệ.'
    else if CompareText(Copy(App, 1, Length(R) + 1), R + '\') <> 0 then
      Result := 'Chế độ thử: /DIR phải nằm trong TESTROOT.'
    else if CompareText(ExtractFileName(App), 'cad-mcp') <> 0 then
      Result := 'Thư mục cài phải có tên "cad-mcp".';
    Exit;
  end;
  { SEC-3: DisableDirPage không chặn /DIR= -> chỉ cho đúng %LOCALAPPDATA%\cad-mcp }
  if CompareText(App, Norm(ExpandConstant('{localappdata}\cad-mcp'))) <> 0 then
  begin
    Result := 'Chỉ cài vào ' + Norm(ExpandConstant('{localappdata}\cad-mcp')) + ' (không hỗ trợ /DIR khác).';
    Exit;
  end;
  while AcadRunning do
  begin
    if WizardSilent then
    begin
      Result := 'AutoCAD hoặc Claude AI cho AutoCAD đang chạy. Lưu bản vẽ, đóng AutoCAD rồi chạy lại Setup.';
      Exit;
    end;
    if MsgBox('AutoCAD hoặc Claude AI cho AutoCAD đang chạy. Hãy LƯU bản vẽ và ĐÓNG AutoCAD, rồi bấm Retry (Thử lại).'#13#10 +
      'Setup không tự đóng AutoCAD.', mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := 'Đã huỷ vì AutoCAD hoặc Claude AI cho AutoCAD đang chạy.';
      Exit;
    end;
  end;
  if RegQueryStringValue(HKCU, 'Software\ClaudeCad', 'Home', Old) and (CompareText(Norm(Old), App) <> 0)
     and DirExists(Old) and not WizardSilent then
    MsgBox('Palette hiện đang dùng thư mục:'#13#10 + Old + #13#10#13#10 +
      'Sau khi cài, palette sẽ dùng bản cài mới tại:'#13#10 + App + #13#10#13#10 +
      'Thư mục cũ không bị sửa hay xoá. (Máy phát triển: chạy lại build.cmd để quay về bản dev.)', mbInformation, MB_OK);
end;

{ JSON thuần ASCII: \ " và ký tự ngoài 32..126 -> \uXXXX. Không có IntToHex thì dùng Format('%.4x', [C]) }
function JsonStr(S: String): String;
var
  I, C: Integer;
begin
  Result := '"';
  for I := 1 to Length(S) do
  begin
    C := Ord(S[I]);
    if S[I] = '\' then Result := Result + '\\'
    else if S[I] = '"' then Result := Result + '\"'
    else if (C < 32) or (C > 126) then Result := Result + '\u' + Lowercase(Format('%.4x', [C]))
    else Result := Result + S[I];
  end;
  Result := Result + '"';
end;

procedure WriteMcpJson;
var
  App, S: String;
begin
  App := ExpandConstant('{app}');
  S := '{"mcpServers": {"autocad": {"command": ' + JsonStr(App + '\python\python.exe') +
       ', "args": ["-I", ' + JsonStr(App + '\server.py') + ']}}}';
  if not SaveStringToFile(App + '\mcp.json', S, False) then
    RaiseException('Không ghi được ' + App + '\mcp.json');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Exe, S: String;
  Rc: Integer;
begin
  if CurStep <> ssPostInstall then Exit;
  WriteMcpJson;
  if IsTest then Exit;
  if WizardIsTaskSelected('claudecli') and ((not WizardSilent) or SilentWantsClaude) then
  begin
    WizardForm.StatusLabel.Caption := 'Đang tải Claude Code từ downloads.claude.ai và kiểm chữ ký...';
    if not RunCliScript('-Mode install', Rc) or (Rc <> 0) then Log('Cai Claude Code khong thanh cong, ma ' + IntToStr(Rc));
  end;
  Exe := ClaudeExe;
  if (Exe <> '') and ClaudeSigned(Exe) then
  begin
    VerifiedClaude := Exe;
    { M1: có hạn 15 s trong claude-cli.ps1; KHÔNG ghi output (SEC-8) }
    LoggedIn := RunCliScript('-Mode status -Path "' + Exe + '"', Rc) and (Rc = 0);
    Exit;
  end;
  if Exe = '' then
    S := 'Chưa cài được Claude Code (không có Internet, bị chặn, hoặc đã bỏ chọn). Palette đã cài xong nhưng cần Claude Code để chạy.'#13#10 +
         'Xem Hướng dẫn sử dụng, mục Xử lý sự cố.'
  else
    S := 'Tìm thấy ' + Exe + ' nhưng chữ ký số không phải của Anthropic, PBC. Setup sẽ không tự chạy tệp này.'#13#10 +
         'Xem Hướng dẫn sử dụng, mục Xử lý sự cố.';
  Log(S);
  if not WizardSilent then MsgBox(S, mbError, MB_OK);
end;

{ M1: trạng thái đăng nhập lấy ở ssPostInstall (claude-cli.ps1 -Mode status, có timeout) — không còn gọi claude.exe ở trang Finished }
function OfferLogin: Boolean;
begin
  Result := not (IsTest or (VerifiedClaude = '') or LoggedIn);
end;

{ bản cài THỬ (pack_setup verify): khoá ClaudeCad-Test\Home trỏ đúng thư mục gỡ — cùng phép dò CurUninstallStepChanged dùng để xoá khoá thử }
function IsTestInstall: Boolean;
var
  V: String;
begin
  Result := RegQueryStringValue(HKCU, 'Software\ClaudeCad-Test', 'Home', V) and (CompareText(Norm(V), Norm(ExpandConstant('{app}'))) = 0);
end;

function InitializeUninstall: Boolean;
begin
  Result := True;
  if UninstallSilent and not IsTestInstall and AcadRunning then
  begin
    Log('Go im lang bi tu choi: AutoCAD (hoac python cua palette) dang chay');
    Result := False;
    Exit;
  end;
  while AcadRunning and not UninstallSilent do
    if MsgBox('AutoCAD hoặc Claude AI cho AutoCAD đang chạy và đang giữ tệp của palette. Lưu bản vẽ, đóng AutoCAD rồi bấm Retry (Thử lại).',
      mbError, MB_RETRYCANCEL) = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
end;

{ Tệp ghim bản vẽ: cad_claude_doc_<pid>.txt (3.0.x) hoặc cad_claude_doc_<pid>_<n>.txt (3.1). CHỈ chữ số và tối đa một '_' giữa hai nhóm số }
function IsPinName(N: String): Boolean;
var
  I, U: Integer;
begin
  Result := (Length(N) > 19) and (CompareText(Copy(N, 1, 15), 'cad_claude_doc_') = 0) and (CompareText(Copy(N, Length(N) - 3, 4), '.txt') = 0);
  U := 0;
  if Result then
    for I := 16 to Length(N) - 4 do
      if N[I] = '_' then
      begin
        U := U + 1;
        if (U > 1) or (I = 16) or (I = Length(N) - 4) then Result := False;
      end
      else if (N[I] < '0') or (N[I] > '9') then Result := False;
end;

procedure DeletePinFiles;
var
  R: TFindRec;
  T: String;
begin
  T := ExpandConstant('{%TEMP}');
  if FindFirst(T + '\cad_claude_doc_*.txt', R) then
  try
    repeat
      if IsPinName(R.Name) then DeleteFile(T + '\' + R.Name);
    until not FindNext(R);
  finally
    FindClose(R);
  end;
end;

{ Q3/HK: xoá mọi cài đặt của palette (Home, Visible, Model, Effort, WithSelection, HotkeyCapture, HotkeyToggle) + tệp tạm phiên,
  CHỈ khi đây là bản cài thật (= %LOCALAPPDATA%\cad-mcp) và Home trỏ nó (hoặc không có Home). Home trỏ bản dev -> không đụng. }
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  App, V: String;
begin
  if CurUninstallStep <> usPostUninstall then Exit;
  App := Norm(ExpandConstant('{app}'));
  if CompareText(App, Norm(ExpandConstant('{localappdata}\cad-mcp'))) = 0 then
    if not RegQueryStringValue(HKCU, 'Software\ClaudeCad', 'Home', V) or (CompareText(Norm(V), App) = 0) then
    begin
      RegDeleteKeyIncludingSubkeys(HKCU, 'Software\ClaudeCad');
      DeleteFile(ExpandConstant('{%TEMP}\cad_claude_session.txt'));
      DeleteFile(ExpandConstant('{%TEMP}\cad_claude_doc.txt'));
      DeletePinFiles;
    end;
  if RegQueryStringValue(HKCU, 'Software\ClaudeCad-Test', 'Home', V) and (CompareText(Norm(V), App) = 0) then
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\ClaudeCad-Test');
end;
