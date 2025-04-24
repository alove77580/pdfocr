#define MyAppName "PDF OCR识别工具"
#define MyAppVersion "1.0"
#define MyAppPublisher "星空：www.itzkb.cn"
#define MyAppExeName "PDF_OCR_Tool.exe"

[Setup]
; 注: AppId的值为唯一标识符
AppId={{8F4E9A1D-7B3C-4E5F-9A2B-6C7D8E9F0A1B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
; 以下行取消注释，以在非管理安装模式下运行（仅为当前用户安装）
;PrivilegesRequired=lowest
OutputBaseFilename=PDF_OCR_Tool_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ShowLanguageDialog=no
SetupIconFile=icon.ico

[Messages]
BeveledLabel=星空软件
ButtonNext=下一步(&N) >
ButtonBack=< 上一步(&B)
ButtonCancel=取消(&C)
ButtonInstall=安装(&I)
ButtonFinish=完成(&F)
SelectDirDesc=选择安装位置
SelectDirLabel3=安装程序将安装 [name] 到下列文件夹。
SelectDirBrowseLabel=点击"下一步"继续。要选择其他文件夹，请点击"浏览"。
WizardSelectDir=选择安装位置
WelcomeLabel1=欢迎使用 [name] 安装向导
WelcomeLabel2=这将在您的计算机上安装 [name/ver]。%n%n建议在继续之前关闭所有其他应用程序。
WizardReady=准备安装
ReadyLabel1=安装程序准备开始安装 [name]。
ReadyLabel2=点击"安装"开始安装。要查看或更改设置，请点击"上一步"。
WizardInstalling=正在安装
StatusLabel=正在安装 [name]，请稍候...
WizardFinished=安装完成
FinishedLabel=安装已完成。程序可以通过选择"运行 [name]"选项立即运行。
FinishedRestartLabel=要完成安装，您需要重新启动计算机。您想现在重新启动吗？
FinishedRestartMessage=要完成安装，您需要重新启动计算机。%n%n安装程序将关闭计算机。您想现在重新启动吗？
UninstallStatusLabel=正在从您的计算机卸载 [name]，请稍候...

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务:"; Flags: unchecked

[Files]
Source: "dist\PDF_OCR_Tool\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\PDF_OCR_Tool\tesseract\*"; DestDir: "{app}\tesseract"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dist\PDF_OCR_Tool\poppler\*"; DestDir: "{app}\poppler"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dist\PDF_OCR_Tool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "运行 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // 检查必要的文件是否存在
    if not FileExists(ExpandConstant('{app}\tesseract\tesseract.exe')) then
    begin
      MsgBox('警告：Tesseract-OCR组件可能未正确安装，程序可能无法正常工作。', mbError, MB_OK);
    end;
    
    if not FileExists(ExpandConstant('{app}\poppler\pdfinfo.exe')) then
    begin
      MsgBox('警告：Poppler组件可能未正确安装，程序可能无法正常工作。', mbError, MB_OK);
    end;
  end;
end; 