program txt2doc;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  ActiveX,
  ComObj,
  iniFiles,
  Variants,
  Windows;

const
  FONT_SIZE = 15;
  FONT_NAME = 'Courier New';
  FONT_SCALING = 100;      // default scaling is 100%, so without scaling
  FONT_BOLD = 0;           // normal font weight by default

  PARAGRAPH_LINESPACING = 1.0;  // lincespacing multiplier 

  PAGESETUP_ORIENTATION = 0;  // vertical orientation by default


  PAGESETUP_TOPMARGIN     = 0.7;  // default margins in centimeters
  PAGESETUP_BOTTOMMARGIN  = 0.7;
  PAGESETUP_LEFTMARGIN    = 0.7;
  PAGESETUP_RIGHTMARGIN   = 0.7;


Var
  inputFileName,
  outputFileName,
  configFileName : String;
  exePath: String;

  oWord, doc: Variant;

  fontSize: Integer;
  fontName: String;
  fontBold: Integer;
  fontScaling: Integer;

  paragraphLineSpacing: Single;
  pageSetupOrientation: integer;

  pageSetupTopMargin: Single;
  pageSetupBottomMargin: Single;
  pageSetupLeftMargin: Single;
  pageSetupRightMargin: Single;


function getParam(i: Integer): String;
begin
  if (i >= 1) and (i <= ParamCount) then begin
    Result := ParamStr(i);
  end else begin
    Result := '';
  end;
end;

procedure readConfigFile(configFile: string);
Var ini: TIniFile;
begin
  ini := TIniFile.Create(configFile);
  try
    fontSize := ini.ReadInteger('Font', 'Size', FONT_SIZE);
    fontName := ini.ReadString('Font', 'Name', FONT_NAME);
    fontBold := ini.ReadInteger('Font', 'Bold', FONT_BOLD);    
    fontScaling := ini.ReadInteger('Font', 'Scaling', FONT_SCALING);

    paragraphLineSpacing := ini.ReadFloat('ParagraphFormat','LineSpacing', PARAGRAPH_LINESPACING);

    pageSetupOrientation := ini.ReadInteger('PageSetup','Orientation', PAGESETUP_ORIENTATION);
    pageSetupTopMargin := ini.ReadFloat('PageSetup','TopMargin', PAGESETUP_TOPMARGIN);
    pageSetupBottomMargin := ini.ReadFloat('PageSetup','BottomMargin', PAGESETUP_BOTTOMMARGIN);
    pageSetupLeftMargin := ini.ReadFloat('PageSetup','LeftMargin', PAGESETUP_LEFTMARGIN);
    pageSetupRightMargin := ini.ReadFloat('PageSetup','RightMargin', PAGESETUP_RIGHTMARGIN);

  finally
    ini.Free;
  end;
end;



procedure parseCommandLineParameters;
var i: Integer;
    s: String;
begin
    i := 1;
    while i <=  ParamCount do begin
        s := LowerCase(ParamStr(i));
        if s = '-c' then begin
               configFileName := getParam( i + 1 );
               Inc(i);
        end;
        if s = '-i' then begin
               inputFileName := getParam( i + 1 );
               inc(i);
        end;
        if s = '-o' then begin
               outputFileName := getParam( i + 1 );
               inc(i);
        end;
        inc(i);
    end;
end;

function getCurDrive: Char;
var
  s1: string;
  s2: Char;
begin
  GetDir(0,s1);
  Writeln('GetDir :'  + s1);
  s2     := s1[1];
  Result := s2;
end;


function makeFullPath(fileName: string): string;
begin
  if (Copy(fileName, 1, 1) = '\') then begin
     fileName := getCurDrive() + ':' + fileName;
     Writeln('Change disk letter :'  + fileName);
  end;


  if (not ((length(fileName) >= 3) and
      ((Copy(fileName, 2, 2) = ':\')))) then begin
     fileName := exePath + fileName;
     Writeln('Relative path :' + fileName);
  end;

  Result := fileName;
end;



function CentimetersToPoints(a: Single): Single;
begin
  Result := 28.35 * a;
end;


function LinesToPoints(a: Single): Single;
begin
  Result:= a * 12;
end;  


function StrAnsiToOem(const aStr : String) : String;
var
  Len : Integer;
begin
  Result := '';
  Len := Length(aStr);
  if Len = 0 then Exit;
  SetLength(Result, Len);
  CharToOemBuff(PChar(aStr), PChar(Result), Len);
end;




procedure showUsage;
begin
    Writeln;
    WriteLn('Plain text files to MS Word document files converter v1.0.0');
    WriteLn('(c) tvolf 2014');
    Writeln;
    Writeln('Usage: ');
    Writeln('txt2doc.exe -i <input_file.txt> [-o <output_file.doc>] [-c <config.ini>]');
    Writeln('where');
    Writeln('   <input_file.txt>  - file name of input plain text file (.txt)');
    Writeln('   <output_file.doc> - file name of output MS Word document file (.doc)');
    Writeln('   <config.ini>      - configuration .ini file');
end;


begin
  ExePath := ExtractFilePath(ParamStr(0));

  if ParamCount = 0 then begin
    showUsage;
    ExitCode := -1;
    Exit;
  end;

  parseCommandLineParameters;
  if inputFileName = '' then begin
    showUsage;
    ExitCode := -1;
    Exit;
  end;

  if outputFileName = '' then begin
    outputFileName := ChangeFileExt(inputFileName, '.doc');
  end;


  inputFileName := makeFullPath(inputFileName);
  outputFileName := makeFullPath(outputFileName);

  if inputFileName = outputFileName then begin
    outputFileName := inputFileName + '.doc';
  end;


  if configFileName <> '' then begin
     configFileName := makeFullPath(configFileName);
  end;


  fontSize := FONT_SIZE;
  readConfigFile(configFileName);

  CoInitialize(Nil);
  ExitCode := 0;

  try
     oWord := CreateOleObject('Word.Application');
     Writeln('Word version ' + oWord.Version + ' found');     
  except
     on E: Exception do begin
//       if not VarIsEmpty(oWord) then begin
           oWord.quit;
//       end;
       WriteLn('Cannot load MS Word: ' + StrAnsiToOem(E.Message));
       ExitCode := -2;
     end;
  end;

  try
      doc := oWord.documents.open( FileName := InputFileName, Format:= 4);
  except
     on E: Exception do begin
       if not VarIsEmpty(oWord) then begin
           oWord.quit;
       end;
       WriteLn('Open Document error: ' + StrAnsiToOem(E.Message));
       ExitCode := -2;
       Exit;
     end;
  end;

  try
      doc.range.font.size := fontSize;
      doc.range.font.name := fontName;
      doc.range.font.scaling := fontScaling;

      doc.range.font.bold := fontBold;

      doc.range.ParagraphFormat.LineSpacingRule := 5; // 5 == wdLineSpaceMultiple
      doc.range.ParagraphFormat.LineSpacing :=
             LinesToPoints(paragraphLineSpacing);

      doc.PageSetup.Orientation := pageSetupOrientation;

      doc.PageSetup.TopMargin := CentimetersToPoints(pageSetupTopMargin);
      doc.PageSetup.BottomMargin := CentimetersToPoints(pageSetupBottomMargin);
      doc.PageSetup.LeftMargin := CentimetersToPoints(pageSetupLeftMargin);
      doc.PageSetup.RightMargin := CentimetersToPoints(pageSetupRightMargin);
   except
     on E: Exception do begin
//       if not VarIsEmpty(oWord) then begin
           oWord.quit;
//       end;
       WriteLn('Setting parameters error: ' + StrAnsiToOem(E.Message));
       ExitCode := -2;
       Exit;
     end;
   end;

   try
      doc.SaveAs(FileName := outputFileName, FileFormat := 0);
   except
     on E: Exception do begin
//       if not VarIsEmpty(oWord) then begin
           oWord.quit;
//       end;
       WriteLn('Save Document error: ' + StrAnsiToOem(E.Message));
       ExitCode := -2;
       Exit;
     end;
   end;
    Writeln('Converting is done!');
     if  not VarIsEmpty(oWord) then begin
         oWord.quit;
     end;

end.
