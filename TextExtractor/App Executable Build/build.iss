[Setup]
AppId={{YOUR-UNIQUE-APP-ID}} 
AppName=Text Extractor
AppVersion=1.0
AppPublisher=Ed's Industries
DefaultDirName={autopf}\TextExtractor
DefaultGroupName=Text Extractor
AllowNoIcons=yes
OutputBaseFilename=TextExtractor
SetupIconFile=icon.ico         
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "addtesspath"; Description: "Add Tesseract to system PATH"; GroupDescription: "Optional Components:"; Flags: unchecked

[Files]
Source: "TextExtractor.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "tesseract_installer.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "icon.ico"; DestDir: "{app}" 

[Icons]
Name: "{group}\Text Extractor"; Filename: "{app}\TextExtractor.exe"; IconFilename: "{app}\icon.ico"; HotKey: "ctrl+alt+t"
Name: "{autodesktop}\Text Extractor"; Filename: "{app}\TextExtractor.exe"; IconFilename: "{app}\icon.ico"; HotKey: "ctrl+alt+t"

[Run]
Filename: "{tmp}\tesseract_installer.exe"; \
    Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /DIR=""{app}\Tesseract-OCR"""; \
    StatusMsg: "Installing Tesseract OCR dependency..."; \
    Flags: shellexec waituntilterminated

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  PathKey, CurrentPath, NewPath: string;
begin
  if CurStep = ssPostInstall then begin
    if IsTaskSelected('addtesspath') then begin
      PathKey := 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment';
      if RegQueryStringValue(HKEY_LOCAL_MACHINE, PathKey, 'Path', CurrentPath) then begin
        NewPath := ExpandConstant('{app}') + '\Tesseract-OCR';
        if Pos(LowerCase(NewPath), LowerCase(CurrentPath)) = 0 then begin
          CurrentPath := CurrentPath + ';' + NewPath;
          RegWriteStringValue(HKEY_LOCAL_MACHINE, PathKey, 'Path', CurrentPath);
        end;
      end;
    end;
  end;
end;
