% HOW TO USE THIS FILE
%   1) Run the file and choose which Export file you will be converting
%   2) Choose the directory in which you will store the output files in
%   Standard format
%   3) After conversion you can run STDForm_to_GLD model converter to convert
%   into a GridLAB-D model
%
%   Written by Duncan Rodriguez for CWRU

clear all;

%CURRENTLY ONLY WORKS WITH SHEET TABULATED *.XLS/*.XLSX FILE
dirFile = 'XLSDir.txt';
if ~exist(dirFile)
    fid = fopen(dirFile,'w');
    fclose(fid);
end
dirFile = fopen('XLSDir.txt','r');
initDir1 = fgetl(dirFile);
PTWExport = fgetl(dirFile);
initDir2 = fgetl(dirFile);
fclose(dirFile);
if ~ischar(initDir1)
    initDir1 = pwd;
end
if ~ischar(initDir2)
    initDir2 = pwd;
end
dir = char(initDir1);
dir2 = initDir2;
keepCurrentSettings = questdlg('Use previous file and settings?',...
    'Confirmation',...
    'Yes','No','File only','Yes');
switch keepCurrentSettings
    case 'Yes'
    case 'No'
    case 'File only'
end
if strcmp(keepCurrentSettings,'No')
    %Select the Export file to use and output directory for Standard files
    % Directory for input files (CSVs)
    try
    [PTWExport,dir,filter] = uigetfile({'*.csv;*.xls;*xlsx','All Tabular Files';'*.*','All Files' },'Select XLS file to convert',initDir1);
    catch
        
    end
    dirFile = fopen('XLSDir.txt','w');
    fprintf(dirFile,'%s',dir);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',PTWExport);
    fprintf(dirFile,'\r\n');
    fclose(dirFile);
end
if strcmp(keepCurrentSettings,'File only') || strcmp(keepCurrentSettings,'No')
    % Directory for output of Standard format CSV files
    dir2 = uigetdir(initDir2,'Select Output file directory');
    dirFile = fopen('XLSDir.txt','w');
    fprintf(dirFile,'%s',dir);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',PTWExport);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',dir2);
    fprintf(dirFile,'\n');
    fclose(dirFile);
end
if strcmp(keepCurrentSettings,'Yes') &&  strcmp(initDir1,pwd)
    disp('Error: There are no files or directories given, you must select new settings.');
    return;
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Convert the chosen file into Standard Format
tic
PTWExport = [dir,PTWExport];
[status,sheets,xlFormat] = xlsfinfo(PTWExport);
% Create the 'Buses' file
if sum(contains(sheets,'Bu_ComponentName'))
    disp('Creating Buses File...');
    currfile = 'Buses';
    [num,Name,raw] = tryxlsread(PTWExport,'Bu_ComponentName');
    Name = Name(3:length(Name),1);
    Name = regexprep(Name,{':','&',',',' '},{'CNXN','and','P','_'});
    Name = strcat(Name,'CNXN1');
    %write the file
    nameCol = {'Name'};
    writeCSV(currfile,Name, dir2,nameCol);
end

% Create the 'Capacitors' file
if sum(contains(sheets,'Fi_ComponentName'))
    disp('Creating Capacitor File...');
    currfile = 'Capacitors';
    [num,Name1,raw] = tryxlsread(PTWExport,'Fi_ComponentName');
    Name1 = Name1(3:length(Name1));
    depth1 = ['A2:A',num2str(length(Name1)+1)];
    [num,busc1,raw] = tryxlsread(PTWExport,'Fi_ConnectedComponent1',depth1);
    for i = 1:length(busc1)
        busc1{i} = regexprep(busc1{i},{'.*(?=:2)'},{[Name1{i}]});
    end
    [~,~,Phases1] = tryxlsread(PTWExport,'Fi_Phase',depth1);
    Phases1 = regexprep(cellfun(@num2str, Phases1, 'UniformOutput', false),{'0','1','2','4','3','5','6','7','48','96','80','112'},{'None','A','B','C','AB','CA','BC','ABC','AB','BC','CA','ABC'});
    [Status1,~,~] = tryxlsread(PTWExport,'Fi_InService',depth1);
    [~,~,type1] = tryxlsread(PTWExport,'Fi_FilterType',depth1);
    [Vreg1,~,~] = tryxlsread(PTWExport,'Fi_SystemNominalVoltage',depth1);
    [seqr,~,~] = tryxlsread(PTWExport,'Fi_Positive Sequence Resistor',depth1);
    [seqt,~,~] = tryxlsread(PTWExport,'Fi_Positive Sequence Reactor',depth1);
    [seqc,~,~] = tryxlsread(PTWExport,'Fi_Positive Sequence Capacitor',depth1);
    for i = 1:length(Name1)
        seq(i,1) = max(max(seqr(i),seqt(i)),seqc(i));
    end
    busc1 = regexprep(busc1,{':2',':3'},{':1',':1'});
    busc1 = regexprep(busc1,{':','&',',',' '},{'CNXN','and','P','_'});
    type = type1;
    type = cellstr(num2str(cell2mat(type)));
    caplist = {Name1,busc1,Phases1,Status1,Vreg1,seq,type};%keep type at end
    listlen = length(caplist);
    [type,I] = sort(caplist{listlen});
    for i = 1:listlen
        caplist{i} = caplist{i}(I,:);
    end
    resistor = cell(1,listlen);
    reactor = cell(1,listlen);
    capacitor = cell(1,listlen);
    for i = 1:length(caplist{listlen})
        if sum(strcmp(caplist{listlen}(i),{'0'}))
            resistor = extractRow(caplist,i,resistor);
        elseif sum(strcmp(caplist{listlen}(i),{'1'}))
            reactor = extractRow(caplist,i,reactor);
        elseif sum(strcmp(caplist{listlen}(i),{'2'}))
            capacitor = extractRow(caplist,i,capacitor);
        end
    end
    %write the file
    if ~isempty(capacitor{1})
        currfile = 'Capacitors';
        capacitorTable = capacitor(1:(listlen-1));
        capacitorCol = {'Name','bus','Phases','Status1','Vreg1','X'};
        writeCSV(currfile,capacitorTable, dir2,capacitorCol);
    end
    if ~isempty(resistor{1})
        currfile = 'Groundings';
        resistorTable = resistor(1:(listlen-1));
        resistorCol = {'Name','bus','Phases','Status1','Vreg1','R'};
        writeCSV(currfile,resistorTable, dir2,resistorCol);
    end
end

% Create the 'CapControls' file
disp('Creating CapControl File...');
currfile = 'CapControls.csv';
%write the file
writeFile = fopen(fullfile(dir2,currfile),'w');
fclose(writeFile);

% Create the 'Generators' file
if sum(contains(sheets,'SG_ComponentName'))
    disp('Creating Generators File...');
    currfile = 'Generators';
    [num,Name,raw] = tryxlsread(PTWExport,'SG_ComponentName');
    Name = Name(3:length(Name));
    Name = regexprep(Name,{':','&',',',' '},{'CNXN','and','P','_'});
    
    [num,bussw1,raw] = tryxlsread(PTWExport,'SG_ConnectedComponent1');
    bussw1 = bussw1(2:length(bussw1));
    bussw1 = regexprep(bussw1,{':','&',',',' '},{'CNXN','and','P','_'});
    [Voltage,~,~] = tryxlsread(PTWExport,'SG_SystemNominalVoltage');
    Voltage = Voltage(1:length(Voltage));
    [Status,~,~] = tryxlsread(PTWExport,'SG_InService');
    Status = Status(1:length(Status));
    [PF,~,~] = tryxlsread(PTWExport,'SG_Rated PF');
    PF = PF(1:length(PF));
    [leadlag,~,~] = tryxlsread(PTWExport,'SG_LeadOrLag');
    leadlag = leadlag(1:length(leadlag));
    [kW,~,~] = tryxlsread(PTWExport,'SG_RatedSize');
    kW = kW(1:length(kW));
    [ratedUnits,~,~] = tryxlsread(PTWExport,'SG_RatedUnits');
    ratedUnits = ratedUnits(1:length(ratedUnits));
    [kW,PF] = tokW(Voltage,PF,leadlag,kW, ratedUnits);
    %write the file
    genesTable = {Name,bussw1,Voltage,Status,kW,PF};
    genesCol = {'Name','bussw1','Voltage','Status','kW','PF'};
    writeCSV(currfile,genesTable, dir2,genesCol);
end

% Create the 'Lines' file
if sum(contains(sheets,'Ca_ComponentName'))
    disp('Creating Lines File...');
    currfile = 'Lines';
    [num,Name,raw] = tryxlsread(PTWExport,'Ca_ComponentName');
    Name = Name(3:length(Name));
    Name = regexprep(Name,{':','&',',',' '},{'CNXN','and','P','_'});
    depth = ['A2:A',num2str(length(Name)+1)];
    [num,busl1,raw] = tryxlsread(PTWExport,'Ca_ConnectedComponent1');
    busl1 = busl1(2:length(busl1));
    for i = 1:length(busl1)
        busl1{i} = regexprep(busl1{i},{'.*(?=:2)','.*(?=:3)'},{[Name{i}],[Name{i}]});
    end
    busl1 = regexprep(busl1,{':2',':3'},{':1',':1'});
    busl1 = regexprep(busl1,{':','&',',',' '},{'CNXN','and','P','_'});
    [~,~,Phases] = tryxlsread(PTWExport,'Ca_Phase');
    Phases = Phases(2:length(Phases));
    Phases = regexprep(cellfun(@num2str, Phases, 'UniformOutput', false),{'0','1','2','4','3','5','6','7','48','96','80','112'},{'None','A','B','C','AB','CA','BC','ABC','AB','BC','CA','ABC'});
    [num,busl2,raw] = tryxlsread(PTWExport,'Ca_ConnectedComponent2');
    busl2 = busl2(2:length(busl2));
    curbus = busl2;
    for i = 1:length(curbus)
        a = Name{i};
        b = curbus{i}(1:length(curbus{i})-2);
        l = max(length(a),length(b));
        sumname = num2str(sum(pad(a,l)+pad(b,l)));
        busl2{i} = regexprep(curbus{i},{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
    end
    busl2 = regexprep(busl2,{':','&',',',' '},{'CNXN','and','P','_'});
    [CondperPhase,~,~] = tryxlsread(PTWExport,'Ca_QtyPerPhase');
    [Length,~,~] = tryxlsread(PTWExport,'Ca_Length');
    Length = Length(1:length(Length));
    %Units are simply assumed to be feet since it is U.S. and not specified
    Units = strread(num2str(ones(1,length(Length))),'%s');
    Units = strrep(Units,'1','ft');
    [~,~,CircularMils] = tryxlsread(PTWExport,'Ca_CircularMils');
    CircularMils = cellfun(@(x) x,CircularMils(2:length(CircularMils)),'UniformOutput',false);
    [Status,~,~] = tryxlsread(PTWExport,'Ca_InService');
    Status = Status(1:length(Status));
    Status = strread(num2str(reshape(Status,[1,length(Status)])),'%s');
    %Status = strrep(Status,{'0','1'},{'open','closed'});
    [Voltage,~,~] = tryxlsread(PTWExport,'Ca_SystemNominalVoltage');
    [Resistance,~,~] = tryxlsread(PTWExport,'Ca_Rpos');
    [Reactance,~,~] = tryxlsread(PTWExport,'Ca_Xpos');
    [Ampacity,~,~] = tryxlsread(PTWExport,'Ca_Ampacity');
    [~,~,Installation] = tryxlsread(PTWExport,'Ca_Installation',depth);
    [CableOD,~,~] = tryxlsread(PTWExport,'Ca_CableOD',depth);
    [ConductorOD,~,~] = tryxlsread(PTWExport,'Ca_ConductorOD',depth);
    [RNeutral,~,~] = tryxlsread(PTWExport,'Ca_R Neutral',depth);
    Config = cell(1,length(Name));
    %write the file
    linesTable ={Name,busl1,Phases,busl2,CondperPhase,Length,Units,CircularMils,Status,Voltage,Resistance,Reactance,Ampacity,Installation,CableOD,ConductorOD,Config};
    linesCol ={'Name','busl1','Phases1','busl2','CondperPhase','Length','Units','CircularMils','Status','Voltage','Resistance','Reactance','Ampacity','Installation','CableOD','ConductorOD','Config'};
    writeCSV(currfile,linesTable, dir2,linesCol);
end

% Create the 'Loads' file
if sum(contains(sheets,'No_ComponentName'))
    disp('Creating Loads File...');
    currfile = 'Loads';
    %Non-Motor Loads
    [num,Name1,raw] = tryxlsread(PTWExport,'No_ComponentName');
    Name1 = Name1(3:length(Name1));
    [num,busl1,raw] = tryxlsread(PTWExport,'No_ConnectedComponent1');
    busl1 = busl1(2:length(busl1));
    for i = 1:length(busl1)
        busl1{i} = regexprep(busl1{i},{'.*(?=:2)'},{[Name1{i}]});
    end
    [~,~,numPhases1] = tryxlsread(PTWExport,'No_Phase');
    numPhases1 = numPhases1(2:length(numPhases1));
    [Voltage1,~,~] = tryxlsread(PTWExport,'No_SystemNominalVoltage');
    Voltage1 = Voltage1(1:length(Voltage1));
    [Status1,~,~] = tryxlsread(PTWExport,'No_Load Factor');
    Status1 = Status1(1:length(Status1));
    Status1 = strread(num2str(reshape(Status1,[1,length(Status1)])),'%s');
    Status1 = strrep(Status1,'1','fixed');
    Status1 = strrep(Status1,'%.0f','variable');
    [Model1,~,~] = tryxlsread(PTWExport,'No_InService');
    Model1 = Model1(1:length(Model1));
    [~,~,Connection1] = tryxlsread(PTWExport,'No_ConnectionType');
    Connection1 = Connection1(2:length(Connection1));
    Connection1 = regexprep(cellfun(@num2str, Connection1, 'UniformOutput', false),{'0','1','2','3','4','5','6'},{'Delta', 'Wye','Wye-Ground','Delta Mid-Winding Grnd','Wye-Capacitor-Delta','ZigZag','ZigZag-Ground'});
    [PF1,~,~] = tryxlsread(PTWExport,'No_PF');
    PF1 = PF1(1:length(PF1));
    [leadlag1,~,~] = tryxlsread(PTWExport,'No_LeadOrLag');
    leadlag1 = leadlag1(1:length(leadlag1));
    [kW1,~,~] = tryxlsread(PTWExport,'No_RatedSize');
    kW1 = kW1(1:length(kW1));
    [kVAR1,~,~] = tryxlsread(PTWExport,'No_RatedSize2');
    kVAR1 = kVAR1(1:length(kVAR1));
    [ratedUnits1,~,~] = tryxlsread(PTWExport,'No_RatedUnits');
    ratedUnits1 = ratedUnits1(1:length(ratedUnits1));
    %Induction Motors
    if sum(contains(sheets,'In_ComponentName'))
        [num,Name2,raw] = tryxlsread(PTWExport,'In_ComponentName');
        Name2 = Name2(3:length(Name2));
        [num,busl2,raw] = tryxlsread(PTWExport,'In_ConnectedComponent1');
        busl2 = busl2(2:length(busl2));
        for i = 1:length(busl2)
            busl2{i} = regexprep(busl2{i},{'.*(?=:2)'},{[Name2{i}]});
        end
        [~,~,numPhases2] = tryxlsread(PTWExport,'In_Phase');
        numPhases2 = numPhases2(2:length(numPhases2));
        [Voltage2,~,~] = tryxlsread(PTWExport,'In_SystemNominalVoltage');
        Voltage2 = Voltage2(1:length(Voltage2));
        [Status2,~,~] = tryxlsread(PTWExport,'In_Load Factor');
        Status2 = Status2(1:length(Status2));
        Status2 = strread(num2str(reshape(Status2,[1,length(Status2)])),'%s');
        Status2 = strrep(Status2,'1','fixed');
        Status2 = strrep(Status2,'%.0f','variable');
        [Model2,~,~] = tryxlsread(PTWExport,'In_InService');
        Model2 = Model2(1:length(Model2));
        [~,~,Connection2] = tryxlsread(PTWExport,'In_ConnectionType');
        Connection2 = Connection2(2:length(Connection2));
        Connection2 = regexprep(cellfun(@num2str, Connection2, 'UniformOutput', false),{'0','1','2','3','4','5','6'},{'Delta', 'Wye','Wye-Ground','Delta Mid-Winding Grnd','Wye-Capacitor-Delta','ZigZag','ZigZag-Ground'});
        [PF2,~,~] = tryxlsread(PTWExport,'In_PF');
        PF2 = PF2(1:length(PF2));
        [leadlag2,~,~] = tryxlsread(PTWExport,'In_LeadOrLag');
        leadlag2 = leadlag2(1:length(leadlag2));
        [kW2mult,~,~] = tryxlsread(PTWExport,'In_NumMotors');
        kW2mult = kW2mult(1:length(kW2mult));
        [kW2,~,~] = tryxlsread(PTWExport,'In_RatedSize');
        kW2 = kW2(1:length(kW2)).*kW2mult;
        [kVAR2,~,~] = tryxlsread(PTWExport,'No_RatedSize2');
        kVAR2 = kVAR2(1:length(kVAR2));
        [ratedUnits2,~,~] = tryxlsread(PTWExport,'In_RatedUnits');
        ratedUnits2 = ratedUnits2(1:length(ratedUnits2));
        [efficiency2,~,~] = tryxlsread(PTWExport,'In_Efficiency');
        efficiency2 = efficiency2(1:length(efficiency2));
    else
        [Name2,numPhases2,busl2,Voltage2,Status2,Model2,Connection2,PF2,leadlag2,kW2,kVAR2,ratedUnits2,efficiency2,Config2] = deal([]);
    end
    %Synchronous Motors
    if sum(contains(sheets,'SM_ComponentName'))
        [num,Name3,raw] = tryxlsread(PTWExport,'SM_ComponentName');
        Name3 = Name3(3:length(Name3));
        [num,busl3,raw] = tryxlsread(PTWExport,'SM_ConnectedComponent1');
        busl3 = busl3(2:length(busl3));
        for i = 1:length(busl3)
            busl3{i} = regexprep(busl3{i},{'.*(?=:2)'},{[Name3{i}]});
        end
        [~,~,numPhases3] = tryxlsread(PTWExport,'SM_Phase');
        numPhases3 = numPhases3(2:length(numPhases3));
        [Voltage3,~,~] = tryxlsread(PTWExport,'SM_SystemNominalVoltage');
        Voltage3 = Voltage3(1:length(Voltage3));
        [Status3,~,~] = tryxlsread(PTWExport,'SM_Load Factor');
        Status3 = Status3(1:length(Status3));
        Status3 = strread(num2str(reshape(Status3,[1,length(Status3)])),'%s');
        Status3 = strrep(Status3,'1','fixed');
        Status3 = strrep(Status3,'%.0f','variable');
        [Model3,~,~] = tryxlsread(PTWExport,'SM_InService');
        Model3 = Model3(1:length(Model3));
        [~,~,Connection3] = tryxlsread(PTWExport,'SM_ConnectionType');
        Connection3 = Connection3(2:length(Connection3));
        Connection3 = regexprep(cellfun(@num2str, Connection3, 'UniformOutput', false),{'0','1','2','3','4','5','6'},{'Delta', 'Wye','Wye-Ground','Delta Mid-Winding Grnd','Wye-Capacitor-Delta','ZigZag','ZigZag-Ground'});
        [PF3,~,~] = tryxlsread(PTWExport,'SM_PF');
        PF3 = PF3(1:length(PF3));
        [leadlag3,~,~] = tryxlsread(PTWExport,'SM_LeadOrLag');
        leadlag3 = leadlag3(1:length(leadlag3));
        [kW3mult,~,~] = tryxlsread(PTWExport,'SM_NumMotors');
        kW3mult = kW3mult(1:length(kW3mult));
        [kW3,~,~] = tryxlsread(PTWExport,'SM_RatedSize');
        kW3 = kW3(1:length(kW3)).*kW3mult;
        [kVAR3,~,~] = tryxlsread(PTWExport,'No_RatedSize2');
        kVAR3 = kVAR3(1:length(kVAR3));
        [ratedUnits3,~,~] = tryxlsread(PTWExport,'SM_RatedUnits');
        ratedUnits3 = ratedUnits3(1:length(ratedUnits3));
        [efficiency3,~,~] = tryxlsread(PTWExport,'SM_Efficiency');
        efficiency3 = efficiency3(1:length(efficiency3));
    else
        [Name3,numPhases3,busl3,Voltage3,Status3,Model3,Connection3,PF3,leadlag3,kW3,kVAR3,ratedUnits3,efficiency3] = deal([]);
    end
    Name = [Name1;Name2;Name3];
    Name = regexprep(Name,{':','&',',',' '},{'CNXN','and','P','_'});
    depth3 = length(Name3)+1;
    buslo = [busl1;busl2;busl3];
    buslo = regexprep(buslo,{':2'},{':1'});
    buslo = regexprep(buslo,{':','&',',',' '},{'CNXN','and','P','_'});
    numPhases = [numPhases1;numPhases2;numPhases3];
    Phases = regexprep(cellfun(@num2str, numPhases, 'UniformOutput', false),{'0','1','2','4','3','5','6','7','48','96','80','112'},{'None','A','B','C','AB','CA','BC','ABC','AB','BC','CA','ABC'});
    numPhases = regexprep(cellfun(@num2str, numPhases, 'UniformOutput', false),{'0','2','1','4','3','5','6','7','48','96','80','112'},{'0','1','1','1','2','2','2','3','2','2','2','3'});
    Voltage = [Voltage1;Voltage2;Voltage3];
    Status = [Status1;Status2;Status3];
    Model = [Model1;Model2;Model3];
    Connection = [Connection1;Connection2;Connection3];
    PF = [PF1;PF2;PF3];
    leadlag = [leadlag1;leadlag2;leadlag3];
    kW = [kW1;kW2;kW3];
    kVAR = [kVAR1;kVAR2;kVAR3];
    ratedUnits = [ratedUnits1;ratedUnits2;ratedUnits3];
    [kW,PF] = tokW(Voltage,PF,leadlag,kW, ratedUnits);
    efficiency = [ones(length(Voltage1),1);efficiency2;efficiency3];
    kW = kW./efficiency;
    %write the file
    loadsTable = {Name,buslo,numPhases,Phases,Voltage,Status,Model,Connection,kW,kVAR,PF};
    loadsCol = {'Name','buslo','numPhases','Phases','Voltage','Status','Model','Connection','kW','kVAR','PF'};
    writeCSV(currfile,loadsTable, dir2,loadsCol);
end

% Create the 'LoadXfmrs' file
disp('Creating LoadXfmrs File...');
currfile = 'LoadXfmrs.csv';
%write the file
writeFile = fopen(fullfile(dir2,currfile),'w');
fclose(writeFile);

% Create the 'Motor Controllers' file
if sum(contains(sheets,'Mo_ComponentName'))
    disp('Creating Motor Controllers File...');
    currfile = 'MTRCTRLs';
    [num,Name,raw] = tryxlsread(PTWExport,'Mo_ComponentName');
    Name = Name(3:length(Name));
    depth1 = ['A2:A',num2str(length(Name)+1)];
    Name = regexprep(Name,{':','&',',',' '},{'CNXN','and','P','_'});
    [num,busmo1,raw] = tryxlsread(PTWExport,'Mo_ConnectedComponent1',depth1);
    busmo1 = regexprep(busmo1,{':','&',',',' '},{'CNXN','and','P','_'});
    [num,busmo2,raw] = tryxlsread(PTWExport,'Mo_ConnectedComponent2',depth1);
    curbus = busmo2;
    for i = 1:length(curbus)
        a = Name{i};
        b = curbus{i}(1:length(curbus{i})-2);
        l = max(length(a),length(b));
        sumname = num2str(sum(pad(a,l)+pad(b,l)));
        busmo2{i} = regexprep(curbus{i},{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
    end
    busmo2 = regexprep(busmo2,{':','&',',',' '},{'CNXN','and','P','_'});
    [Voltage,~,~] = tryxlsread(PTWExport,'Mo_SystemNominalVoltage',depth1);
    [PF,~,~] = tryxlsread(PTWExport,'Mo_PF',depth1);
    [power,~,~] = tryxlsread(PTWExport,'Mo_RatedSize',depth1);
    power = power*1000;
    [~,~,ratedUnits] = tryxlsread(PTWExport,'Mo_RatedUnits',depth1);
    ratedUnits = str2num(cell2mat(regexprep(cellfun(@num2str, ratedUnits, 'UniformOutput', false),{'0','1'},{'4','3'})));
    leadlag = zeros(length(Name),1);
    [kW,PF] = tokW(Voltage,PF,leadlag,power,ratedUnits);
    [efficiency,~,~] = tryxlsread(PTWExport,'Mo_Efficiency',depth1);
    %write the file
    MCTable = {Name,busmo1,busmo2,Voltage,power,PF,efficiency};
    MCCol = {'Name','busmo1','busmo2','Voltage','power','PF','efficiency'};
    writeCSV(currfile,MCTable, dir2,MCCol);
end

% Create the 'Regulators' file
if sum(contains(sheets,'Pr_ComponentName'))
    disp('Creating Regulators File...');
    [num,Name1,raw] = tryxlsread(PTWExport,'Pr_ComponentName');
    Name1 = Name1(3:length(Name1));
    depth1 = ['A2:A',num2str(length(Name1)+1)];
    [num,busr11,raw] = tryxlsread(PTWExport,'Pr_ConnectedComponent1',depth1);
    for i = 1:length(busr11)
        busr11{i} = regexprep(busr11{i},{'.*(?=:2)','.*(?=:3)'},{[Name1{i}],[Name1{i}]});
    end
    [~,~,Phases1] = tryxlsread(PTWExport,'Pr_Poles',depth1);
    Phases1 = regexprep(cellfun(@num2str, Phases1, 'UniformOutput', false),{'0','1'},{'ABC','A'});
    [num,busr21,raw] = tryxlsread(PTWExport,'Pr_ConnectedComponent2',depth1);
    curbus = busr21;
    for i = 1:length(curbus)
        a = Name1{i};
        b = curbus{i}(1:length(curbus{i})-2);
        l = max(length(a),length(b));
        sumname = num2str(sum(pad(a,l)+pad(b,l)));
        busr21{i} = regexprep(curbus{i},{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
    end
    [Status1,~,~] = tryxlsread(PTWExport,'PR_InService',depth1);
    [Vreg1,~,~] = tryxlsread(PTWExport,'Pr_SystemNominalVoltage',depth1);
    [~,~,primary] = tryxlsread(PTWExport,'Pr_CT Primary',depth1);
    primary = str2double(regexprep(cellfun(@num2str, primary, 'UniformOutput', false),'NaN','0'));
    [~,~,secondary] = tryxlsread(PTWExport,'Pr_CT Secondary',depth1);
    secondary = str2double(regexprep(cellfun(@num2str, secondary, 'UniformOutput', false),'NaN','0'));
    PTratio1 = num2cell(secondary./primary);
    [raw,txt,SensorTrip1] = tryxlsread(PTWExport,'Pr_Sensor-Trip',depth1);
    [~,~,CTrating1] = tryxlsread(PTWExport,'Pr_Mom_Rating',depth1);
    [~,~,Band1] = tryxlsread(PTWExport,'Pr_VoltageRangeFactor K',depth1);
    [~,~,RR1] = tryxlsread(PTWExport,'Pr_Test X-R',depth1);
    XX1 = regexprep(cellfun(@num2str, RR1, 'UniformOutput', false),{'[^0]','\w*1'},{'1','1'});%%%%%%%%%%
    BasekVA1 = num2cell(Vreg1.*primary);
    [~,~,type1] = tryxlsread(PTWExport,'Pr_ProtectionType');
    type1 = type1(2:length(type1));
    %Pi Equivalent
    if sum(contains(sheets,'Pi_ComponentName'))
        [num,Name2,raw] = tryxlsread(PTWExport,'Pi_ComponentName');
        Name2 = Name2(3:length(Name2));
        depth2 = ['A2:A',num2str(length(Name2)+1)];
        [num,busr12,raw] = tryxlsread(PTWExport,'Pi_ConnectedComponent1',depth2);
        for i = 1:length(busr12)
            busr12{i} = regexprep(busr12{i},{'.*(?=:2)','.*(?=:3)'},{[Name2{i}],[Name2{i}]});
        end
        [~,~,Phases2] = tryxlsread(PTWExport,'Pi_Phase',depth2);
        Phases2 = regexprep(cellfun(@num2str, Phases2, 'UniformOutput', false),{'0','1','2','4','3','5','6','7','48','96','80','112'},{'None','A','B','C','AB','CA','BC','ABC','AB','BC','CA','ABC'});
        [num,busr22,raw] = tryxlsread(PTWExport,'Pi_ConnectedComponent2',depth2);
        curbus = busr22;
        for i = 1:length(curbus)
            a = Name2{i};
            b = curbus{i}(1:length(curbus{i})-2);
            l = max(length(a),length(b));
            sumname = num2str(sum(pad(a,l)+pad(b,l)));
            busr22{i} = regexprep(curbus{i},{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
        end
        [Status2,~,~] = tryxlsread(PTWExport,'Pi_InService',depth2);
        [Vreg2,~,~] = tryxlsread(PTWExport,'Pi_SystemNominalVoltage',depth2);
        PTratio2 = cellstr(num2str(zeros(length(Vreg2),1)));
        CTrating2 = PTratio2;
        Band2 = PTratio2;
        [raw,txt,SensorTrip2] = tryxlsread(PTWExport,'Pi_Ampacity',depth2);
        [~,~,RR2] = tryxlsread(PTWExport,'Pi_Rpos',depth2);
        [~,~,XX2] = tryxlsread(PTWExport,'Pi_Xpos',depth2);
        [~,~,BasekVA2] = tryxlsread(PTWExport,'Pi_Base kVA',depth2);
        [~,~,type2] = tryxlsread(PTWExport,'Pi_PiType',depth2);
    else
        [Name2,busr12,Phases2,busr22,Status2,Vreg2,SensorTrip2,PTratio2,CTrating2,Band2,RR2,XX2,BasekVA2,type2] = deal([]);
    end
    Name = [Name1;Name2];
    Name = regexprep(Name,{':','&',',',' '},{'CNXN','and','P','_'});
    busr1 = [busr11;busr12];
    busr1 = regexprep(busr1,{':2',':3'},{':1',':1'});
    %bus1Para = regexp(bus1,{'-\S',',\S'},{'match','match'});
    %bus1Para = [bus1Para{:}];
    %bus1Para = regexprep(bus1Para,{'-',','},{'_','_'});
    busr1 = regexprep(busr1,{':','&',',',' '},{'CNXN','and','P','_'});
    Phases = [Phases1;Phases2];
    busr2 = [busr21;busr22];
    busr2 = regexprep(busr2,{':','&',',',' '},{'CNXN','and','P','_'});
    Status = [Status1;Status2];
    Vreg = [Vreg1;Vreg2];
    SensorTrip = [SensorTrip1;SensorTrip2];
    PTratio = [PTratio1;PTratio2];
    PTratio = regexprep(cellfun(@num2str, PTratio, 'UniformOutput', false),'NaN','0');
    CTrating = [CTrating1;CTrating2];
    CTrating = regexprep(cellfun(@num2str, CTrating, 'UniformOutput', false),'NaN','0');
    Band = [Band1;Band2];
    Band = regexprep(cellfun(@num2str, Band, 'UniformOutput', false),'NaN','0');
    RR = [RR1;RR2];
    XX = [XX1;XX2];
    BasekVA = [BasekVA1;BasekVA2];
    type = [type1;type2];
    type  = [num2cell(zeros(length(XX)-length(type),1));type];
    type = cellstr(num2str(cell2mat(type)));
    reglist = {Name,busr1,Phases,busr2,Status,Phases,Vreg,SensorTrip,PTratio,CTrating,Band,RR,XX,BasekVA,type};%keep type at end
    listlen = length(reglist);
    [type,I] = sort(reglist{listlen});
    for i = 1:listlen
        reglist{i} = reglist{i}(I,:);
    end
    breakers = cell(1,listlen);
    fuses = cell(1,listlen);
    relays = cell(1,listlen);
    switches = cell(1,listlen);
    HVMVs = cell(1,listlen);
    MTRPros = cell(1,listlen);
    Reactors = cell(1,listlen);
    regs = cell(1,listlen);
    len1 = 5;
    lenV = 7;
    lenSTrip = 8;
    for i = 1:length(reglist{listlen})
        if sum(strcmp(reglist{listlen}(i),{'    1','11001','11002','11003','11004'}))
            breakers = extractRow(reglist,i,breakers);
        elseif sum(strcmp(reglist{listlen}(i),{'11041','11042'}))
            fuses = extractRow(reglist,i,fuses);
        elseif sum(strcmp(reglist{listlen}(i),{'11061','11062','11063','11101','11102'}))
            relays = extractRow(reglist,i,relays);
        elseif sum(strcmp(reglist{listlen}(i),{'11114'}))
            switches = extractRow(reglist,i,switches);
        elseif sum(strcmp(reglist{listlen}(i),{'11081','11082'}))
            HVMVs = extractRow(reglist,i,HVMVs);
        elseif sum(strcmp(reglist{listlen}(i),{'11021'}))
            MTRPros = extractRow(reglist,i,MTRPros);
        elseif sum(strcmp(reglist{listlen}(i),{'    2'}))
            Reactors = extractRow(reglist,i,Reactors);
        else
            regs = extractRow(reglist,i,regs);
        end
    end
    %write the file
    if ~isempty(breakers{1})
        currfile = 'Breakers';
        breakerTable = breakers([1:(len1),lenV,lenSTrip]);
        breakerCol = {'Name','bus1','phases','bus2','status','Voltage','SensorTrip'};
        writeCSV(currfile,breakerTable, dir2,breakerCol);
    end
    if ~isempty(fuses{1})
        currfile = 'Fuses';
        fuseTable = fuses([1:(len1),lenV,lenSTrip]);
        fuseCol = {'Name','busx1','phases','busx2','status','Voltage','SensorTrip'};
        writeCSV(currfile,fuseTable, dir2,fuseCol);
    end
    if ~isempty(relays{1})
        currfile = 'Relays';
        relayTable = relays([1:(len1),lenV]);
        relayCol = {'Name','busx1','phases','busx2','status','Voltage'};
        writeCSV(currfile,relayTable, dir2,relayCol);
    end
    
    switch1Table = switches([1,2,2,4:lenV]);
    switch1Table{3} = cell(length(switch1Table{1}),1);
    
    if ~isempty(HVMVs{1})
        currfile = 'HVMVs';
        HVMVTable = HVMVs([1:(len1),lenV]);
        HVMVCol = {'Name','busx1','phases','busx2','status','Voltage'};
        writeCSV(currfile,HVMVTable, dir2,HVMVCol);
    end
    if ~isempty(MTRPros{1})
        currfile = 'MTRPros';
        MTRProTable = MTRPros([1:(len1),lenV]);
        MTRProCol = {'Name','busx1','phases','busx2','status','Voltage'};
        writeCSV(currfile,MTRProTable, dir2,MTRProCol);
    end
    if ~isempty(Reactors{1})
        currfile = 'Reactors';
        ReactorsTable = Reactors([1:(len1),12,13]);
        ReactorsCol = {'Name','busx1','phases','busx2','status','R','X'};
        writeCSV(currfile,ReactorsTable, dir2,ReactorsCol);
    end
    if ~isempty(regs{1})
        currfile = 'Regulators';
        regTable = regs(1:(listlen-1));
        regCol = {'Name','busr1','Phases','busr2','status','Phases','Vreg','SensorTrip','PTratio','CTrating','Band','RR','XX','BasekVA'};
        writeCSV(currfile,regTable, dir2,regCol);
    end
end

% Create the 'Switches' file
if sum(contains(sheets,'Au_ComponentName'))
    disp('Creating Switches File...');
    currfile = 'Switches';
    [num,Namesw,raw] = tryxlsread(PTWExport,'Au_ComponentName');
    Namesw = Namesw(3:length(Namesw));
    Namesw = [Namesw;switch1Table{1}];
    Namesw = regexprep(Namesw,{':','&',',',' '},{'CNXN','and','P','_'});
    [num,bussw1,raw] = tryxlsread(PTWExport,'Au_ConnectedComponent1');
    bussw1 = bussw1(2:length(bussw1));
    for i = 1:length(bussw1)
        bussw1{i} = regexprep(bussw1{i},{'.*(?=:2)'},{[Namesw{i}]});
    end
    bussw1 = [bussw1;switch1Table{2}];
    bussw1 = regexprep(bussw1,{':2'},{':1'});
    bussw1 = regexprep(bussw1,{':','&',',',' '},{'CNXN','and','P','_'});
    [num,bussw2,raw] = tryxlsread(PTWExport,'Au_ConnectedComponent2');
    bussw2 = bussw2(2:length(bussw2));
    for i = 1:length(bussw2)
        bussw2{i} = regexprep(bussw2{i},{'.*(?=:2)'},{[Namesw{i}]});
    end
    bussw2 = [bussw2;switch1Table{3}];
    bussw2 = regexprep(bussw2,{':2'},{':1'});
    bussw2 = regexprep(bussw2,{':','&',',',' '},{'CNXN','and','P','_'});
    [num,bussw3,raw] = tryxlsread(PTWExport,'Au_ConnectedComponent3');
    bussw3 = bussw3(2:length(bussw3));
    curbus = bussw3;
    for i = 1:length(curbus)
        a = Namesw{i};
        b = curbus{i}(1:length(curbus{i})-2);
        l = max(length(a),length(b));
        sumname = num2str(sum(pad(a,l)+pad(b,l)));
        bussw3{i} = regexprep(curbus{i},{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
    end
    bussw3 = [bussw3;switch1Table{4}];
    bussw3 = regexprep(bussw3,{':','&',',',' '},{'CNXN','and','P','_'});
    [SwitchPosition,~,~] = tryxlsread(PTWExport,'Au_Switch Position');
    SwitchPosition = SwitchPosition(1:length(SwitchPosition));
    SwitchPosition = [SwitchPosition;switch1Table{5}];
    %write the file
    switchTable = {Namesw,bussw1,bussw2,bussw3,SwitchPosition};
    switchCol = {'Namesw','bussw1','bussw2','bussw3','SwitchPosition'};
    writeCSV(currfile,switchTable, dir2,switchCol);
end

% Create the 'Transformers' file
if sum(contains(sheets,'2-_ComponentName'))
    disp('Creating Transformers File...');
    currfile = 'Transformers';
    [num,Name1,raw] = tryxlsread(PTWExport,'2-_ComponentName');
    Name1 = Name1(3:length(Name1));
    depth1 = ['A2:A',num2str(length(Name1)+1)];
    [~,~,phases1] = tryxlsread(PTWExport,'2-_Phase',depth1);
    phases1 = regexprep(cellfun(@num2str, phases1, 'UniformOutput', false),{'0','1','2','4','3','5','6','7','48','96','80','112'},{'None','A','B','C','AB','CA','BC','ABC','ABN','BCN','CAN','ABCN'});
    [~,~,busx11] = tryxlsread(PTWExport,'2-_ConnectedComponent1',depth1);
    for i = 1:length(busx11)
        busx11{i} = regexprep(busx11{i},{'.*(?=:2)'},{[Name1{i}]});
    end
    [~,~,busx21] = tryxlsread(PTWExport,'2-_ConnectedComponent2',depth1);
    curbus = busx21;
    for i = 1:length(curbus)
        a = Name1{i};
        b = curbus{i}(1:length(curbus{i})-2);
        l = max(length(a),length(b));
        sumname = num2str(sum(pad(a,l)+pad(b,l)));
        busx21{i} = regexprep(curbus{i},{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
    end
    [kV_pri1,~,~] = tryxlsread(PTWExport,'2-_Pri RatedVoltage',depth1);
    [kV_sec1,~,~] = tryxlsread(PTWExport,'2-_Sec RatedVoltage',depth1);
    [MVA1,~,~] = tryxlsread(PTWExport,'2-_Nominal kVA',depth1);
    [~,~,Conn_pri1] = tryxlsread(PTWExport,'2-_Pri Connection',depth1);
    [~,~,Conn_sec1] = tryxlsread(PTWExport,'2-_Sec Connection',depth1);
    [XHL1,~,~] = tryxlsread(PTWExport,'2-_Xpos',depth1);
    [RHL1,~,~] = tryxlsread(PTWExport,'2-_Rpos',depth1);
    %3 Winding Transformers
    if sum(contains(sheets,'3-_ComponentName'))
        [num,Name2,raw] = tryxlsread(PTWExport,'3-_ComponentName');
        Name2 = Name2(3:length(Name2));
        Name3 = strcat(Name2,'2');
        Name2 = strcat(Name2,'1');
        depth2 = ['A2:A',num2str(length(Name2)+1)];
        Name2 = [Name2;Name3];
        phases2 = regexprep(Name2,{'\S*'},{'ABC'});
        [~,~,busx12] = tryxlsread(PTWExport,'3-_ConnectedComponent1',depth2);
        for i = 1:length(busx12)
            busx12{i} = regexprep(char(busx12{i}),{'.*(?=:2)'},{[Name2{i}]});
        end
        [~,~,busx22] = tryxlsread(PTWExport,'3-_ConnectedComponent2',depth2);
        curbus = busx22;
        for i = 1:length(curbus)
            a = Name2{i};
            b = char(curbus{i}(1:length(char(curbus{i}))-2));
            l = max(length(a),length(b));
            sumname = num2str(sum(pad(a,l)+pad(b,l)));
            busx22{i} = regexprep(char(curbus{i}),{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
        end
        [~,~,busx23] = tryxlsread(PTWExport,'3-_ConnectedComponent3',depth2);
        curbus = busx23;
        for i = 1:length(curbus)
            a = Name2{i};
            b = char(curbus{i}(1:length(char(curbus{i}))-2));
            l = max(length(a),length(b));
            sumname = num2str(sum(pad(a,l)+pad(b,l)));
            busx23{i} = regexprep(char(curbus{i}),{'.*(?=:2)','.*(?=:3)'},{['Node',sumname],['Node',sumname]});
        end
        busx31 = regexprep(busx21,{' ','\S'},{'_',''});
        [kV_pri2,~,~] = tryxlsread(PTWExport,'3-_Pri RatedVoltage',depth2);
        [kV_sec2,~,~] = tryxlsread(PTWExport,'3-_Sec RatedVoltage',depth2);
        [kV_sec3,~,~] = tryxlsread(PTWExport,'3-_Ter RatedVoltage',depth2);
        [MVA2,~,~] = tryxlsread(PTWExport,'3-_Pri-Sec Base kVA',depth2);
        [MVA3,~,~] = tryxlsread(PTWExport,'3-_Pri-Ter Base kVA',depth2);
        [~,~,Conn_pri2] = tryxlsread(PTWExport,'3-_Pri Connection',depth2);
        [~,~,Conn_sec2] = tryxlsread(PTWExport,'3-_Sec Connection',depth2);
        [~,~,Conn_sec3] = tryxlsread(PTWExport,'3-_Ter Connection',depth2);
        [XHL2,~,~] = tryxlsread(PTWExport,'3-_Pri-Sec Xpos',depth2);
        [XHL3,~,~] = tryxlsread(PTWExport,'3-_Pri-Ter Xpos',depth2);
        [RHL2,~,~] = tryxlsread(PTWExport,'3-_Pri-Sec Rpos',depth2);
        [RHL3,~,~] = tryxlsread(PTWExport,'3-_Pri-Ter Rpos',depth2);
    else
        [Name2,phases2,busx12,busx22,busx23,kV_pri2,kV_sec2,kV_sec3,MVA2,MVA3,Conn_pri2,Conn_sec2,Conn_sec3,XHL2,XHL3,RHL2,RHL3] = deal([]);
    end
    Name = [Name1;Name2;];
    Name = regexprep(Name,{':','&',',',' '},{'CNXN','and','P','_'});
    phases = [phases1;phases2];
    busx1 = [busx11;busx12;busx12];
    busx1 = regexprep(busx1,{':2'},{':1'});
    busx1 = regexprep(busx1,{':','&',',',' '},{'CNXN','and','P','_'});
    busx2 = [busx21;busx22;busx23];
    busx2 = regexprep(busx2,{':','&',',',' '},{'CNXN','and','P','_'});
    kV_pri = {[kV_pri1;kV_pri2;kV_pri2]};
    kV_pri = cell2mat(cellfun(@(x) x/1000,kV_pri,'UniformOutput',false));
    kV_sec = {[kV_sec1;kV_sec2;kV_sec3]};
    kV_sec = cell2mat(cellfun(@(x) x/1000,kV_sec(1:length(kV_sec)),'UniformOutput',false));
    MVA = {[MVA1;MVA2;MVA3]};
    MVA = cell2mat(cellfun(@(x) x/1000,MVA(1:length(MVA)),'UniformOutput',false));
    Conn_pri = [Conn_pri1;Conn_pri2;Conn_pri2];
    Conn_pri = regexprep(cellfun(@num2str, Conn_pri, 'UniformOutput', false),{'0','1','2','3','4','5','6'},{'DELTA', 'WYE','GWYE','Delta Mid-Winding Grnd','Wye-Capacitor-Delta','ZigZag','ZigZag-Ground'});
    Conn_sec = [Conn_sec1;Conn_sec2;Conn_sec3];
    Conn_sec = regexprep(cellfun(@num2str, Conn_sec, 'UniformOutput', false),{'0','1','2','3','4','5','6'},{'DELTA', 'WYE','GWYE','Delta Mid-Winding Grnd','Wye-Capacitor-Delta','ZigZag','ZigZag-Ground'});
    XHL = {[XHL1;XHL2;XHL3]};
    XHL = cell2mat(cellfun(@(x) x/100,XHL,'UniformOutput',false)); %percent to per-unit
    RHL = {[RHL1;RHL2;RHL3]};
    RHL = cell2mat(cellfun(@(x) x/100,RHL,'UniformOutput',false)); %percent to per-unit
    %write the file
    xfmrTable = {Name,busx1,phases,busx2,kV_pri,kV_sec,MVA,Conn_pri,Conn_sec,XHL,RHL};
    xfmrCol = {'Name','busx1','phases','busx2','kV_pri','kV_sec','MVA','Conn_pri','Conn_sec','XHL','RHL'};
    writeCSV(currfile,xfmrTable,dir2,xfmrCol);
end

% Create the 'Triplex_Lines' file
disp('Creating Triplex Lines File...');
currfile = 'Triplex_Lines.csv';
%write the file
writeFile = fopen(fullfile(dir2,currfile),'w');
fclose(writeFile);

% Create the 'UnbalancedLoads' file
disp('Creating Unbalanced Loads File...');
currfile = 'UnbalancedLoads.csv';
%write the file
writeFile = fopen(fullfile(dir2,currfile),'w');
fclose(writeFile);

% Create the 'Utilities' file
if sum(contains(sheets,'Ut_ComponentName'))
    disp('Creating Utilities File...');
    currfile = 'Utilities';
    [num,Nameut,raw] = tryxlsread(PTWExport,'Ut_ComponentName');
    Nameut = Nameut(3:length(Nameut));
    depth1 = ['A2:A',num2str(length(Nameut)+1)];
    [num,busut,raw] = tryxlsread(PTWExport,'Ut_ConnectedComponent1',depth1);
    busut = regexprep(busut,{':','&',',',' '},{'CNXN','and','P','_'});
    [~,~,phases] = tryxlsread(PTWExport,'2-_Phase',depth1);
    phases = regexprep(cellfun(@num2str, phases, 'UniformOutput', false),{'0','1','2','4','3','5','6','7','48','96','80','112'},{'None','A','B','C','AB','CA','BC','ABC','ABN','BCN','CAN','ABCN'});
    [Voltage,~,~] = tryxlsread(PTWExport,'Ut_SystemNominalVoltage',depth1);
    utilTable = {Nameut,busut,phases,Voltage};
    utilCol = {'Name','busut','phases','Voltage'};
    writeCSV(currfile,utilTable, dir2,utilCol);
end
disp(['Elapsed Time: ' num2str(toc)]);