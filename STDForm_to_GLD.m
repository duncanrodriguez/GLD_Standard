clear all;
format long g;
convertAlone = "Yes";
tic
%% Most of the things you might want to change via a scripting mechanism are located in this section
if convertAlone == "Yes"
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
        %Select the Export file to use and output directory for Jason Fuller files
        % Directory for input files (CSVs)
        [PTWExport,dir,filter] = uigetfile({'*.csv;*.xls;*xlsx','All Tabular Files';'*.*','All Files' },'Select XLS file to convert',initDir1);
        dirFile = fopen('XLSDir.txt','w');
        fprintf(dirFile,'%s',dir);
        fprintf(dirFile,'\r\n');
        fprintf(dirFile,'%s',PTWExport);
        fprintf(dirFile,'\r\n');
        fclose(dirFile);
    end
    if strcmp(keepCurrentSettings,'File only') || strcmp(keepCurrentSettings,'No')
        % Directory for output of Jason Fuller format CSV files
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
end

tic
dirFile = fopen('XLSDir.txt','r');
initDir1 = fgetl(dirFile);
if strcmp(initDir1,pwd)|| ~ischar(initDir1)
    disp('Error: There are no files or directories given, you must select new settings.');
    return;
end
fileName = fgetl(dirFile);
fileName = regexprep(fileName,{'(?=xls).*','\.'},{'',''});
initDir2 = fgetl(dirFile);
fclose(dirFile);
% Directory for input files (CSVs)
dir = initDir2;
dir2 = dir;

% Power flow solver method
solver_method = 'NR';

% Start and stop times
start_date='''2000-09-01';
stop_date  = '''2000-09-01';
start_time='12:00:00''';
stop_time = '17:00:00''';
timezone='PST+8PDT';

%Are you using CSV headers?
containsHeaders = 'n';

% Voltage regulator and capacitor settings
%  All voltages in on 120 volt "per unit" basis
%  VAr setpoints for capacitors are in kVAr
%  Time is in seconds

% Regulator bandcenter voltage, bandwidth voltage, time delay
reg = [7500/60, 2,  60;  % VREG1 (at feeder head)
    7470/60, 2, 120;  % VREG2 (cascaded reg on north side branch, furthest down circuit)
    7440/60, 2,  75;  % VREG3 (cascaded reg on north side branch, about halfway up circuit before VREG2)
    7380/60, 2,  90]; % VREG4 (solo reg on south side branch)

% Capacitor voltage high, voltage low, kVAr high, kVAr low, time delay
% - Note, Cap0-Cap2 are in VOLTVAR control mode, Cap3 is in MANUAL mode
% -- (Cap3 is on south side branch after VREG 4)
cap = [130, 115.5, 375, -250, 480;  % CapBank0 (right before VREG2, but after VREG3)
    130, 115.5, 325, -250, 300;  % CapBank1 (a little after substation, before VREG3 or VREG4))
    130, 115.5, 350, -250, 180]; % CapBank2 (at substation)



%% Some nominal voltage stuff for assigning flat start voltages
nom_volt1 = '7199.558';
nom_volt2 = '12470.00';
nom_volt3 = '69715.05';
nom_volt4 = '115000.00';


%% Load Breakers.csv values
% Name1|bus2|phases3|bus4|Status5|Voltage6|IRating
fidBreakers = fopen([dir '\Breakers.csv']);
RawBreakers = [];
if fidBreakers ~= -1
    if strcmp(containsHeaders,'y')
        Header1Breakers = textscan(fidBreakers,'%s',1);
    end
    Header2Breakers = textscan(fidBreakers,'%s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawBreakers = textscan(fidBreakers,'%s %s %s %s %n %n %n','Delimiter',',');
end

% Load Buses.csv values
% Name1
fidBuses = fopen([dir '\Buses.csv']);
RawBuses = [];
if fidBuses ~= -1
    if strcmp(containsHeaders,'y')
        Header1Buses = textscan(fidBuses,'%s',1);
    end
    Header2Buses = textscan(fidBuses,'%s',1,'Delimiter',',');
    
    RawBuses = textscan(fidBuses,'%s','Delimiter',',');
end

% Load Capacitors.csv values
% Name1|bus2|phases3|Status4|Voltage5|Impedance6
fidCapacitors = fopen([dir '\Capacitors.csv']);
RawCaps = [];
if fidCapacitors ~= -1
    if strcmp(containsHeaders,'y')
        Header1Capacitors = textscan(fidCapacitors,'%s',1);
    end
    Header2Capacitors = textscan(fidCapacitors,'%s %s %s %s %s %s',1,'Delimiter',',');
    
    RawCaps= textscan(fidCapacitors,'%s %s %s %n %n %n','Delimiter',',');
end

% Load Fuses.csv values
% Name1|bus2|phases3|bus4|Status5|Voltage6|IRating
fidFuses = fopen([dir '\Fuses.csv']);
RawFuses = [];
if fidFuses ~= -1
    if strcmp(containsHeaders,'y')
        Header1Fuses = textscan(fidFuses,'%s',1);
    end
    Header2Fuses = textscan(fidFuses,'%s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawFuses = textscan(fidFuses,'%s %s %s %s %n %n %n','Delimiter',',');
end

% Load Generators.csv values
% Name1|Bus2|Voltage3|Status4|PF5
fidGenes = fopen([dir '\Generators.csv']);
RawGenes = [];
if fidGenes ~= -1
    if strcmp(containsHeaders,'y')
        Header1Genes = textscan(fidGenes,'%s',1);
    end
    Header2Genes = textscan(fidGenes,'%s %s %s %s %s %s',1,'Delimiter',',');
    
    RawGenes = textscan(fidGenes,'%s %s %n %n %n %n','Delimiter',',');
end

% Load Grounding.csv values
% Name1|bus2|phases3|Status4|Voltage5|R6
fidGnds = fopen([dir '\Groundings.csv']);
RawGnds = [];
if fidGnds ~= -1
    if strcmp(containsHeaders,'y')
        Header1Gnds = textscan(fidGnds,'%s',1);
    end
    Header2Gnds = textscan(fidGnds,'%s %s %s %s %s %s',1,'Delimiter',',');
    
    RawGnds = textscan(fidGnds,'%s %s %s %n %n %n','Delimiter',',');
end

% Load HVMVs.csv values
% Name1|bus2|phases3|bus4|Status5|Voltage6|IRating
fidHVMVs = fopen([dir '\HVMVs.csv']);
RawHVMVs = [];
if fidHVMVs ~= -1
    if strcmp(containsHeaders,'y')
        Header1HVMVs = textscan(fidHVMVs,'%s',1);
    end
    Header2HVMVs = textscan(fidHVMVs,'%s %s %s %s %s %s',1,'Delimiter',',');
    
    RawHVMVs = textscan(fidHVMVs,'%s %s %s %s %n %n','Delimiter',',');
end

% Load Loads.csv values
% Name1|bus2|numberphases3|Phases4|Voltage5|Type6|Status7|Connection8|kW9|kVAR10|PF11
fidLoads = fopen([dir '\Loads.csv']);
RawLoads = [];
if fidLoads ~= -1
    if strcmp(containsHeaders,'y')
        Header1Loads = textscan(fidLoads,'%s',1);
    end
    Header2Loads = textscan(fidLoads,'%s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawLoads = textscan(fidLoads,'%s %s %n %s %n %s %n %s %n %n %n','Delimiter',',');
end

% Load Lines.csv values

% Name1|From node2|Phases3|to
% node4|CondperPhase5|Length6|Units7|CircularMils8|Status9|Voltage10|Resistance11|Reactance12|Ampacity13|Installation14|Cable
% OD15|Conductor OD16|Config17
fidLines = fopen([dir '\Lines.csv']);
RawLines = [];
if fidLines ~= -1
    if strcmp(containsHeaders,'y')
        Header1Lines = textscan(fidLines,'%s',1);
    end
    Header2Lines = textscan(fidLines,'%s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawLines = textscan(fidLines,'%s %s %s %s %n %n %s %n %n %n %n %n %n %s %n %n %s','Delimiter',',');
end

% Load MTRPros.csv values
% Name1|bus2|phases3|bus4|Status5||Voltage6
fidMTRPros = fopen([dir '\MTRPros.csv']);
RawMTRPros = [];
if fidMTRPros ~= -1
    if strcmp(containsHeaders,'y')
        Header1MTRPros = textscan(fidMTRPros,'%s',1);
    end
    Header2MTRPros = textscan(fidMTRPros,'%s %s %s %s %s %s',1,'Delimiter',',');
    
    RawMTRPros = textscan(fidMTRPros,'%s %s %s %s %n %n','Delimiter',',');
end

% Load MTRCTRLs.csv values
% Name1|From node2|to node3|Voltage4|Power5|PF6|Efficiency7
fidMTRCTRLs = fopen([dir '\MTRCTRLs.csv']);
RawMTRCTRLs = [];
if fidMTRCTRLs ~= -1
    if strcmp(containsHeaders,'y')
        Header1MTRCTRLs = textscan(fidMTRCTRLs,'%s',1);
    end
    Header2MTRCTRLs = textscan(fidMTRCTRLs,'%s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawMTRCTRLs = textscan(fidMTRCTRLs,'%s %s %s %n %n %n %n','Delimiter',',');
end

% Load Reactors.csv values
% Name1|From node2|Phases3|to node4|Status5|R6|X7
fidReactors = fopen([dir '\Reactors.csv']);
RawReactors = [];
if fidReactors ~= -1
    if strcmp(containsHeaders,'y')
        Header1Reactors = textscan(fidReactors,'%s',1);
    end
    Header2Reactors = textscan(fidReactors,'%s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawReactors = textscan(fidReactors,'%s %s %s %s %n %n %n','Delimiter',',');
end

% Load Relays.csv values
% Name1|bus2|phases3|bus4|Status5|Voltage6
fidRelays = fopen([dir '\Relays.csv']);
RawRelays = [];
if fidRelays ~= -1
    if strcmp(containsHeaders,'y')
        Header1Relays = textscan(fidRelays,'%s',1);
    end
    Header2Relays = textscan(fidRelays,'%s %s %s %s %s %s',1,'Delimiter',',');
    
    RawRelays = textscan(fidRelays,'%s %s %s %s %n %n','Delimiter',',');
end

% Load Switches.csv values
% Name1|From node1 2|From node2 3|to node4|Switch Position5
fidSwitches = fopen([dir '\Switches.csv']);
RawSwitches = [];
if fidSwitches ~= -1
    if strcmp(containsHeaders,'y')
        Header1Switches = textscan(fidSwitches,'%s',1);
    end
    Header2Switches = textscan(fidSwitches,'%s %s %s %s %s',1,'Delimiter',',');
    
    RawSwitches = textscan(fidSwitches,'%s %s %s %s %n','Delimiter',',');
end

% Load Transformers.csv values
% Name1|From2|phases3|To4|primV5|secV6|MVA7|PrimConn8|SecConn9|%X10|%R11
fidTrans = fopen([dir '\Transformers.csv']);
RawTrans = [];
if fidTrans ~= -1
    if strcmp(containsHeaders,'y')
        Header1Trans = textscan(fidTrans,'%s',1);
    end
    Header2Trans = textscan(fidTrans,'%s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawTrans = textscan(fidTrans,'%s %s %s %s %n %n %n %s %s %n %n','Delimiter',',');
end

% Load LoadXfmrs.csv values
% Name1|#ofPhases2|From3|Phase4|PrimkV5|PrimkVA6|ToPh1-7|Ph1-8|SeckVPh1-9|
% SeckVAPh1-10|ToPh2-11|Ph2-12|SeckVPh2-13|SeckVAPh2-14|imag-15|R1-16|
% R2-17|R3-18|NoLoad-19|X12-20|X13-21|X23-22
fidLoadTrans = fopen([dir '\LoadXfmrs.csv']);
RawLoadTrans = [];
if fidLoadTrans ~= -1
    if strcmp(containsHeaders,'y')
        Header1LoadTrans = textscan(fidLoadTrans,'%s',1);
    end
    Header2LoadTrans = textscan(fidLoadTrans,'%s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawLoadTrans = textscan(fidLoadTrans,'%s %n %s %s %n %n %s %s %n %n %s %s %n %n %n %n %n %n %n %n %n %n','Delimiter',',');
end

% Load Triplex_Lines.csv values
% Name1|From2|Phases3|To4|Phases5|LineConf6|Length7|Units8
fidTripLines = fopen([dir '\Triplex_Lines.csv']);
RawTripLines = [];
if fidTripLines ~= -1
    if strcmp(containsHeaders,'y')
        Header1TripLines = textscan(fidTripLines,'%s',14);
        Header2TripLines = textscan(fidTripLines,'%s',10);
        Header3TripLines = textscan(fidTripLines,'%s',16);
    end
    Header4TripLines = textscan(fidTripLines,'%s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawTripLines = textscan(fidTripLines,'%s %s %s %s %s %s %n %s','Delimiter',',');
end

% Load Loads.csv values
% Name1|bus2|numberphases3|Phases4|Voltage5|Type6|Status7|Connection8|kW9|kVAR10|PF11
fidTripLoads = fopen([dir '\Loads.csv']);
RawTripLoads = [];
if fidTripLoads ~= -1
    if strcmp(containsHeaders,'y')
        Header1TripLoads = textscan(fidTripLoads,'%s',12);
        Header2TripLoads = textscan(fidTripLoads,'%s',8);
        Header3TripLoads = textscan(fidTripLoads,'%s',11);
        Header4TripLoads = textscan(fidTripLoads,'%s',10);
    end
    Header5TripLoads = textscan(fidTripLoads,'%s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawTripLoads = textscan(fidTripLoads,'%s %s %n %s %n %s %n %s %n %n %n','Delimiter',',');
end

% Name1|Bus2|Phase3|Bus4|Phase5|Vreg6|PTratio7|CTrating8|Band9|R10|X11|BasekVA12
fidReg = fopen([dir '\Regulators.csv']);
RawReg = [];
if fidReg ~= -1
    if strcmp(containsHeaders,'y')
        Header1Reg = textscan(fidReg,'%s',1);
    end
    Header2Reg = textscan(fidReg,'%s %s %s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawReg = textscan(fidReg,'%s %s %s %s %n %n %n %n %n %n %n %n %n','Delimiter',',');
end

% Name1
fidUtil = fopen([dir '\Utilities.csv']);
RawUtil = [];
if fidUtil ~= -1
    if strcmp(containsHeaders,'y')
        Header1Util = textscan(fidUtil,'%s',1);
    end
    Header2Util = textscan(fidUtil,'%s %s %s %s',1,'Delimiter',',');
    
    RawUtil = textscan(fidUtil,'%s %s %s %n','Delimiter',',');
end


%Configurations files
% Name1|Acond2|Bcond3|Ccond4|Ncond5|Spacing6
fidConfig = fopen([dir '\Config.csv']);
RawConfig = [];
if fidConfig ~= -1
    if strcmp(containsHeaders,'y')
        Header1Config = textscan(fidConfig,'%s',1);
    end
    Header2Config = textscan(fidConfig,'%s %s %s %s %s %s',1,'Delimiter',',');
    
    RawConfig = textscan(fidConfig,'%s %s %s %s %s %s','Delimiter',',');
end


% Name1|Type|Material2|Strands3|Resistance4|Cable OD5|Conductor OD|Neutral D|GMR6|Ampacity7|Shield OD8|Shield
% Thickness9
fidCondConfig = fopen([dir '\ConductorConfig.csv']);
RawCondConfig = [];
if fidCondConfig ~= -1
    if strcmp(containsHeaders,'y')
        Header1CondConfig = textscan(CondConfig,'%s',1);
    end
    Header2CondConfig = textscan(fidCondConfig,'%s %s %s %s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawCondConfig = textscan(fidCondConfig,'%s %s %s %n %n %n %n %n %n %n %n %n %n %n','Delimiter',',');
end


% Name1|Acond2|Bcond3|Ccond4|Ncond5|Spacing6
fidSpacing = fopen([dir '\Spacing.csv']);
RawSpacing = [];
if fidSpacing ~= -1
    if strcmp(containsHeaders,'y')
        Header1Spacing = textscan(fidSpacing,'%s',1);
    end
    Header2Spacing = textscan(fidSpacing,'%s %s %s %s %s %s %s %s %s %s %s',1,'Delimiter',',');
    
    RawSpacing = textscan(fidSpacing,'%s %n %n %n %n %n %n %n %n %n %n','Delimiter',',');
end


% Values{1}-name | {2}-ohms/km | {3}-GMR in cm | {4}-outer rad? (cm)

Racunits = 'Ohm/km';
GMRunits = 'cm';

fclose('all');

NameLines = char(RawLines{1});
FromLines = char(RawLines{2});
PhasesLines = char(RawLines{3});
ToLines = char(RawLines{4});
LengthLines = (RawLines{6});
UnitLines = char(RawLines{7});
ConfigLines = char(RawLines{8});
StatusLines = char(RawLines{9});

EndBreakers = numrows(RawBreakers);
EndBuses = numrows(RawBuses);
EndCaps = numrows(RawCaps);
EndCondConfig = numrows(RawCondConfig);
EndConfig = numrows(RawConfig);
EndFuses = numrows(RawFuses);
EndHVMVs = numrows(RawHVMVs);
EndGenes = numrows(RawGenes);
EndGnds = numrows(RawGnds);
EndLoads = numrows(RawLoads);
EndLines = numrows(RawLines);
EndMTRCTRLs = numrows(RawMTRCTRLs);
EndMTRPros = numrows(RawMTRPros);
EndReactors = numrows(RawReactors);
EndRelays = numrows(RawRelays);
EndSpacing = numrows(RawSpacing);
EndSwitches = numrows(RawSwitches);
EndTrans = numrows(RawTrans);
EndLoadTrans = numrows(RawLoadTrans);
EndTripLines = numrows(RawTripLines);
EndTripLoads = numrows(RawTripLoads);
EndTripNodes = numrows(RawTripLines);
EndRegs = numrows(RawReg);
EndUtil = numrows(RawUtil);

%% Print to glm file
if strcmp(solver_method,'FBS')
        open_name = [dir2 ['\',fileName,'_FBS.glm']];
elseif strcmp(solver_method,'NR')
        open_name = [dir2 ['\',fileName,'.glm']];
else
    fprintf('screw-up in naming of open file');
end

fid = fopen(open_name,'wt');

%% Header stuff and schedules
fprintf(fid,'//%s system.\n',fileName);
fprintf(fid,'//  Generated %s using Matlab %s.\n\n',datestr(clock),version);

fprintf(fid,'clock {\n');
fprintf(fid,'     timezone %s;\n',timezone);
fprintf(fid,'     starttime %s %s;\n',start_date,start_time);
fprintf(fid,'     stoptime %s %s;\n',stop_date,stop_time);
fprintf(fid,'}\n\n');


%%
fprintf(fid,'module powerflow {\n');
fprintf(fid,'    solver_method %s;\n',solver_method);
fprintf(fid,'    line_limits FALSE;\n');
fprintf(fid,'    line_capacitance TRUE;\n');
fprintf(fid,'    default_maximum_voltage_error 1e-4;\n');
fprintf(fid,'};\n');
fprintf(fid,'module tape;\n\n');
recorder = 'recorders';
fprintf(fid,'#include "%s1.glm";\n',recorder);
%fprintf(fid,'#include "schedules.glm";\n\n');
fprintf(fid,'#set profiler=1;\n');
fprintf(fid,'#set relax_naming_rules=1;\n');
fprintf(fid,'#set suppress_repeat_messages=1;\n');
fprintf(fid,'#set savefile="%s.xml";\n',open_name);
fprintf(fid,'object fault_check{\n');
fprintf(fid,'check_mode SINGLE;\n');
fprintf(fid,'output_filename whatiswrong.txt;\n');
fprintf(fid,'}\n');
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Regulator objects -- Easiest by hand
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

fprintf(fid,'// Regulators and regulator configurations\n\n');

for i=1:EndRegs
    if RawReg{5}(i)
        Config = 'Config';
        %             if ~isempty(RawReg{14}(i))
        %                 Config = char(RawReg{14}(i));
        %             end
        fprintf(fid,'object regulator {\n');
        fprintf(fid,['     name "',char(RawReg{1}(i)),'";\n']);
        fprintf(fid,'     phases %s;\n',char(RawReg{3}(i)));
        fprintf(fid,['     from "',char(RawReg{2}(i)),'";\n']);
        fprintf(fid,['     to "',char(RawReg{4}(i)),'";\n']);
        fprintf(fid,'     configuration %s%s;\n',char(RawReg{1}(i)),Config);
        fprintf(fid,'}\n\n');
        fprintf(fid,'object regulator_configuration {\n');
        fprintf(fid,'     connect_type 1;\n');
        fprintf(fid,'     name %s%s;\n',char(RawReg{1}(i)),Config);
        fprintf(fid,'     Control OUTPUT_VOLTAGE;\n');
        fprintf(fid,'     band_center %.2f;\n',RawReg{6}(i));
        fprintf(fid,'     band_width %.2f;\n',RawReg{10}(i));
        fprintf(fid,'     current_transducer_ratio %.1f;\n',RawReg{9}(i));
        fprintf(fid,'     power_transducer_ratio %.1f;\n',RawReg{8}(i));
        fprintf(fid,'     compensator_r_setting_A %.1f;\n',RawReg{11}(i));
        fprintf(fid,'     compensator_x_setting_A %.1f;\n',RawReg{12}(i));
        fprintf(fid,'     compensator_r_setting_B %.1f;\n',RawReg{11}(i));
        fprintf(fid,'     compensator_x_setting_B %.1f;\n',RawReg{12}(i));
        fprintf(fid,'     compensator_r_setting_C %.1f;\n',RawReg{11}(i));
        fprintf(fid,'     compensator_x_setting_C %.1f;\n',RawReg{12}(i));
        fprintf(fid,'     time_delay %.1f;\n',reg(1,3));
        fprintf(fid,'     raise_taps 16;\n');
        fprintf(fid,'     lower_taps 16;\n');
        fprintf(fid,'     regulation 0.1;\n');
        fprintf(fid,'     Type B;\n');
        fprintf(fid,'}\n\n');
    end
end
for i=1:EndReactors
    if RawReactors{5}(i)
        fprintf(fid,'object series_reactor {\n');
        fprintf(fid,['     name "',char(RawReactors{1}(i)),'";\n']);
    fprintf(fid,'     phases %s;\n',char(RawReactors{3}(i)));
    fprintf(fid,['     from "',char(RawReactors{2}(i)),'";\n']);
    fprintf(fid,['     to "',char(RawReactors{4}(i)),'";\n']);
    fprintf(fid,'     phases %s;\n',char(RawReactors{3}(i)));
    %base = (RawReg{6}(i)/1000)^2/(RawReg{13}(i)/1000);
    base = 1;
    real = base*RawReactors{6}(i);
    imag = base*RawReactors{7}(i);
    fprintf(fid,'     phase_A_impedance %.4f+%.4fj;\n',real,imag);
    fprintf(fid,'     phase_B_impedance %.4f+%.4fj;\n',real,imag);
    fprintf(fid,'     phase_C_impedance %.4f+%.4fj;\n',real,imag);
    fprintf(fid,'}\n\n');
    end
end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Capacitor objects
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

fprintf(fid,'// Capacitors\n\n');
for i = 1:EndCaps
    if RawCaps{4}(i) == 1
        var = (RawCaps{5}(i))^2/RawCaps{6}(i);
        fprintf(fid,'object capacitor {\n');
        fprintf(fid,['     phases ',char(RawCaps{3}(i)),';\n']);
        fprintf(fid,'     name %s;\n',char(RawCaps{1}(i)));
        fprintf(fid,'     pt_phase "%s";\n',char(RawCaps{3}(i)));
        fprintf(fid,'     parent "%s";\n',char(RawCaps{2}(i)));
        fprintf(fid,'     remote_sense "%s";\n',char(RawCaps{2}(i)));
        fprintf(fid,'     phases_connected "%s";\n',char(RawCaps{3}(i)));
        fprintf(fid,'     control MANUAL;\n');
        fprintf(fid,'     VAr_set_high "%.1f";\n',0.15*var);
        fprintf(fid,'     VAr_set_low "%.1f";\n',0.1*var);
        if contains(char(RawCaps{3}(i)),'A')
            fprintf(fid,'     capacitor_A "%.1f";\n',var);
            fprintf(fid,'     switchA CLOSED;\n');
        else
            fprintf(fid,'     switchA OPEN;\n');
        end
        if contains(char(RawCaps{3}(i)),'B')
            fprintf(fid,'     capacitor_B "%.1f";\n',var);
            fprintf(fid,'     switchB CLOSED;\n');
        else
            fprintf(fid,'     switchB OPEN;\n');
        end
        if contains(char(RawCaps{3}(i)),'C')
            fprintf(fid,'     capacitor_C "%.1f";\n',var);
            fprintf(fid,'     switchC CLOSED;\n');
        else
            fprintf(fid,'     switchC OPEN;\n');
        end
        fprintf(fid,'     control_level INDIVIDUAL;\n');
        fprintf(fid,'     time_delay 2.0;\n');
        fprintf(fid,'     cap_nominal_voltage %.1f;\n',0.95*RawCaps{5}(i));
        fprintf(fid,'     nominal_voltage %.1f;\n',RawCaps{5}(i));
        fprintf(fid,'}\n\n');
    end
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Groundings.\n');
disp('Printing Groundings...');
for i = 1:EndGnds
    if RawGnds{4}(i) == 1
        r = RawGnds{6}(i)*3;
        x = 0;
        fprintf(fid,'object line_configuration {\n');
        fprintf(fid,'     name %sConfig;\n',char(RawGnds{1}(i)));
        fprintf(fid,'     z11 %.2f+0j;\n',r);
        fprintf(fid,'     z12 0.01+0.01j;\n');
        fprintf(fid,'     z13 0.01+0.01j;\n');
        fprintf(fid,'     z21 0.01+0.01j;\n');
        fprintf(fid,'     z22 %.2f+0j;\n',r);
        fprintf(fid,'     z23 0.01+0.01j;\n');
        fprintf(fid,'     z31 0.01+0.01j;\n');
        fprintf(fid,'     z32 0.01+0.01j;\n');
        fprintf(fid,'     z33 %.2f+0j;\n',r);
        fprintf(fid,'}\n\n');
        fprintf(fid,'object overhead_line {\n');
        fprintf(fid,'     name "%s";\n',char(RawGnds{1}(i)));
        fprintf(fid,'     phases ABCN;\n');
        fprintf(fid,'     from "%s";\n',char(RawGnds{2}(i)));
        fprintf(fid,'     to "NodeGND";\n');
        fprintf(fid,'     length 5280;\n');
        fprintf(fid,'     configuration %sConfig;\n',char(RawGnds{1}(i)));
        fprintf(fid,'}\n\n');
        %         fprintf(fid,'object load {\n');
        %         fprintf(fid,'     name "%s";\n',char(RawGnds{1}(i)));
        %         fprintf(fid,'     parent "%s";\n',char(RawGnds{3}(i)));
        %         fprintf(fid,'     nominal_voltage %7.0f;\n',RawGnds{5}(i));
        %         fprintf(fid,'     phases %s;\n',char(RawGnds{2}(i)));
        %         fprintf(fid,'     constant_impedance_A %.1f+%.1fj;\n',r,x);
        %         fprintf(fid,'     constant_impedance_B %.1f+%.1fj;\n',r,x);
        %         fprintf(fid,'     constant_impedance_C %.1f+%.1fj;\n',r,x);
        %         fprintf(fid,'}\n\n');
    end
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Transformer objects
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

fprintf(fid,'// Transformer and configuration at feeder\n\n');

for i = 1:EndTrans
    fprintf(fid,'object transformer_configuration {\n');
    from = char(RawTrans{2}(i));
    to = char(RawTrans{4}(i));
    VPri = RawTrans{5}(i);
    VSec = RawTrans{6}(i);
    if (containsRelaxed(char(RawTrans{8}(i)),'DELTA')&& containsRelaxed(char(RawTrans{9}(i)),'WYE'))||(containsRelaxed(char(RawTrans{8}(i)),'DELTA')&& containsRelaxed(char(RawTrans{9}(i)),'GWYE'))
        fprintf(fid,'     connect_type DELTA_GWYE;\n');
    elseif (containsRelaxed(char(RawTrans{8}(i)),'GWYE')&& containsRelaxed(char(RawTrans{9}(i)),'DELTA'))||(containsRelaxed(char(RawTrans{8}(i)),'WYE')&& containsRelaxed(char(RawTrans{9}(i)),'DELTA'))
        fprintf(fid,'     connect_type DELTA_GWYE;\n');
        from = char(RawTrans{4}(i));
        to = char(RawTrans{2}(i));
        VPri = RawTrans{6}(i);
        VSec = RawTrans{5}(i);
    elseif(containsRelaxed(char(RawTrans{8}(i)),'GWYE')&& containsRelaxed(char(RawTrans{9}(i)),'GWYE'))||(containsRelaxed(char(RawTrans{8}(i)),'GWYE')&& containsRelaxed(char(RawTrans{9}(i)),'WYE'))||(containsRelaxed(char(RawTrans{8}(i)),'WYE')&& containsRelaxed(char(RawTrans{9}(i)),'GWYE'))
        fprintf(fid,'     connect_type WYE_WYE;\n');
    else
        fprintf(fid,'     connect_type %s_%s;\n',char(RawTrans{8}(i)),char(RawTrans{9}(i)));
    end
    fprintf(fid,['     name ',char(RawTrans{1}(i)),'Config;\n']);
    fprintf(fid,'     install_type PADMOUNT;\n');
    fprintf(fid,'     power_rating %5.0fkVA;\n',1000*RawTrans{7}(i));
    fprintf(fid,'     primary_voltage %3.2fkV;\n',VPri);
    fprintf(fid,'     secondary_voltage %2.2fkV;\n',VSec);
    fprintf(fid,'     reactance %1.6f;\n',1*RawTrans{10}(i));
    fprintf(fid,'     resistance %1.6f;\n',1*RawTrans{11}(i));
    fprintf(fid,'}\n\n');
    
    fprintf(fid,'object transformer {\n');
    fprintf(fid,['     phases ',char(RawTrans{3}(i)),';\n']);
    fprintf(fid,'     name "%s";\n',char(RawTrans{1}(i)));
    fprintf(fid,'     from "%s";\n',from);
    fprintf(fid,'     to "%s";\n',to);
    fprintf(fid,['     configuration ',char(RawTrans{1}(i)),'Config;\n']);
    fprintf(fid,'}\n\n');
end


%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Center-tap Transformer objects
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

fprintf(fid,'// Center-tap transformer configurations\n\n');

RHL = 0.006;
RHT = 0.012;
RLT = 0.012;

XHL = 0.0204;
XHT = 0.0204;
XLT = 0.0136;

XH = 0.5*(XHL+XHT-XLT);
XL = 0.5*(XHL+XLT-XHT);
XT = 0.5*(XLT+XHT-XHL);

for i=1:EndLoadTrans
    t_conf = sprintf('%.0f%.0f%s',RawLoadTrans{6}(i),RawLoadTrans{10}(i),char(RawLoadTrans{4}(i)));
    t_confs(i,1:length(t_conf)) = t_conf;
    if i==1
        fprintf(fid,'object transformer_configuration {\n');
        fprintf(fid,'     name %s;\n',t_conf);
        fprintf(fid,'     connect_type SINGLE_PHASE_CENTER_TAPPED;\n');
        fprintf(fid,'     install_type POLETOP;\n');
        fprintf(fid,'     primary_voltage %5.1fV;\n',1000*RawLoadTrans{5}(i));
        fprintf(fid,'     secondary_voltage %3.1fV;\n',1000*RawLoadTrans{9}(i));
        fprintf(fid,'     power_rating %2.1fkVA;\n',RawLoadTrans{6}(i));
        fprintf(fid,'     power%s_rating %2.1fkVA;\n',char(RawLoadTrans{4}(i)),RawLoadTrans{10}(i));
        fprintf(fid,'     impedance %f+%fj;\n',RHL,XH);
        fprintf(fid,'     impedance1 %f+%fj;\n',RHT,XL);
        fprintf(fid,'     impedance2 %f+%fj;\n',RLT,XT);
        Z = 7200^2 / (1000 * RawLoadTrans{6}(i) * 0.005);
        R = 7200^2 / (1000 * RawLoadTrans{6}(i) * 0.002);
        fprintf(fid,'     shunt_impedance %.0f+%.0fj;\n',R,Z);
        fprintf(fid,'}\n\n');
    else
        stop = 0;
        for m=1:(i-1)
            if (strcmp(t_conf(1:length(t_conf)),t_confs(m,1:length(t_conf))))
                stop = 1;
                m = i-2;
            end
        end
        
        if stop ~= 1
            fprintf(fid,'object transformer_configuration {\n');
            fprintf(fid,'     name %s;\n',t_conf);
            fprintf(fid,'     connect_type SINGLE_PHASE_CENTER_TAPPED;\n');
            fprintf(fid,'     install_type POLETOP;\n');
            fprintf(fid,'     primary_voltage %5.1f;\n',1000*RawLoadTrans{5}(i));
            fprintf(fid,'     secondary_voltage %3.1f;\n',1000*RawLoadTrans{9}(i));
            fprintf(fid,'     power_rating %2.1f;\n',RawLoadTrans{6}(i));
            fprintf(fid,'     power%s_rating %2.1f;\n',char(RawLoadTrans{4}(i)),RawLoadTrans{10}(i));
            fprintf(fid,'     impedance 0.006+0.0136j;\n');
            fprintf(fid,'     impedance1 0.012+0.0204j;\n');
            fprintf(fid,'     impedance2 0.012+0.0204j;\n');
            Z = 7200^2 / (1000 * RawLoadTrans{6}(i) * 0.005);
            R = 7200^2 / (1000 * RawLoadTrans{6}(i) * 0.002);
            fprintf(fid,'     shunt_impedance %.0f+%.0fj;\n',R,Z);
            fprintf(fid,'}\n\n');
        end
    end
    
    
end

fprintf(fid,'// Center-tap transformers\n\n');

for i=1:EndLoadTrans
    fprintf(fid,'object transformer {\n');
    fprintf(fid,'     configuration "%.0f%.0f%s";\n',RawLoadTrans{6}(i),RawLoadTrans{10}(i),char(RawLoadTrans{4}(i)));
    fprintf(fid,'     name "%s";\n',char(RawLoadTrans{1}(i)));
    fprintf(fid,'     from "%s";\n',char(RawLoadTrans{3}(i)));
    fprintf(fid,'     to "%s";\n',char(RawLoadTrans{7}(i)));
    fprintf(fid,'     nominal_voltage %s;\n',nom_volt1);
    fprintf(fid,'     phases %sS;\n',char(RawLoadTrans{4}(i)));
    fprintf(fid,'}\n\n');
end



%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Triplex-Load objects
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Triplex Node Objects with loads\n\n');
    for i=1:EndTripLoads
        if RawTripLoads{7}(i)
            numPhases = RawTripLoads{3}(i);
            fprintf(fid,'object load {\n');
            fprintf(fid,'     name "%s";\n',char(RawTripLoads{1}(i)));
            fprintf(fid,'     parent "%s";\n',char(RawTripLoads{2}(i)));
            fprintf(fid,'     nominal_voltage %7.0f;\n',RawTripLoads{5}(i));
            fprintf(fid,'     phases %s;\n',char(RawTripLoads{4}(i)));
            if ~isnan(RawTripLoads{9}(i))
                reload = RawTripLoads{9}(i)*1000/numPhases;
                if ~isnan(RawTripLoads{11}(i))
                    imload = RawTripLoads{9}(i)*1000*tan(acos(RawTripLoads{11}(i)))/numPhases;
                else
                    if ~isnan(RawTripLoads{10}(i))
                        imload = RawTripLoads{10}(i)*1000/numPhases;
                    else
                        disp(['Warning: There is not enough load information for',char(RawTripLoads{1}(i)),'.'])
                    end
                end
            else
                if ~isnan(RawTripLoads{10}(i))
                    imload = RawTripLoads{10}(i)*1000/numPhases;
                    reload = RawTripLoads{10}(i)*1000*(1-tan(acos(RawTripLoads{11}(i))))/numPhases;
                else
                    disp(['Warning: There is not enough load information for',char(RawTripLoads{1}(i)),'.'])
                end
            end
            if containsRelaxed(char(RawTripLoads{4}(i)),'A')
            fprintf(fid,'     constant_power_A %.1f+%.1fj;\n',reload,imload);
            end
            if containsRelaxed(char(RawTripLoads{4}(i)),'B')
            fprintf(fid,'     constant_power_B %.1f+%.1fj;\n',reload,imload);
            end
            if containsRelaxed(char(RawTripLoads{4}(i)),'C')
            fprintf(fid,'     constant_power_C %.1f+%.1fj;\n',reload,imload);
            end
            fprintf(fid,'}\n\n');
        end
    end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Triplex-Node objects (non-load objects)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

fprintf(fid,'// Triplex Node Objects without loads\n\n');
disp('Printing triplex nodes...');
for i=1:EndTripNodes
    fprintf(fid,'object triplex_node {\n');
    fprintf(fid,'     name "%s";\n',char(RawTripLines{2}(i)));
    fprintf(fid,'     nominal_voltage 120;\n');
    TphN = char(RawTripLines{4}(i));
    PhNode = TphN(10);
    fprintf(fid,'     phases %sS;\n',PhNode);
    fprintf(fid,'}\n\n');
end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Node objects
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
disp('Printing nodes...');
fprintf(fid,'// Node Objects\n\n');

%1) Name    2) Voltage    3)Connection to or from    4)Phases
% Go through 'From' node list
n=0;
for i=1:EndLines
    if RawLines{9}(i)
        n=n+1;
        Node_Name{1}(n) = RawLines{2}(i);
        Node_Name{2}(n) = RawLines{10}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawLines{3}(i);
    end
end
%To of Breakers
for i=(1):(EndBreakers)
    if RawBreakers{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawBreakers{4}(i);
        Node_Name{2}(n) = RawBreakers{6}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawBreakers{3}(i);
    end
end

%To of Fuses
for i=(1):(EndFuses)
    if RawFuses{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawFuses{4}(i);
        Node_Name{2}(n) = RawFuses{6}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawFuses{3}(i);
    end
end
%To of generators
for i=(1):(EndGenes)
    n=n+1;
    Node_Name{1}(n) = RawGenes{2}(i);
    Node_Name{2}(n) = RawGenes{3}(i);
    Node_Name{3}(n) = 2;
    Node_Name{4}(n) = RawGenes{2}(i);
end
%To of HVMVs
for i=(1):(EndHVMVs)
    if RawHVMVs{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawHVMVs{4}(i);
        Node_Name{2}(n) = RawHVMVs{6}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawHVMVs{3}(i);
    end
end
%To of MTRCTRLs
for i=(1):(EndMTRCTRLs)
    n=n+1;
    Node_Name{1}(n) = RawMTRCTRLs{3}(i);
    Node_Name{2}(n) = RawMTRCTRLs{4}(i);
    Node_Name{3}(n) = 2;
    Node_Name{4}(n) = {'ABC'};
end
%To of MTRPros
for i=(1):(EndMTRPros)
    if RawMTRPros{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawMTRPros{4}(i);
        Node_Name{2}(n) = RawMTRPros{6}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawMTRPros{3}(i);
    end
end
%To of Relays
for i=(1):(EndRelays)
    if RawRelays{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawRelays{4}(i);
        Node_Name{2}(n) = RawRelays{6}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawRelays{3}(i);
    end
end
%From of MTRCTRLs
for i=(1):(EndMTRCTRLs)
    n=n+1;
    Node_Name{1}(n) = RawMTRCTRLs{2}(i);
    Node_Name{2}(n) = RawMTRCTRLs{4}(i);
    Node_Name{3}(n) = 1;
    Node_Name{4}(n) = {'ABC'};
end
%From of MTRPros
for i=(1):(EndMTRPros)
    if RawMTRPros{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawMTRPros{2}(i);
        Node_Name{2}(n) = RawMTRPros{6}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawMTRPros{3}(i);
    end
end
%From of Relays
for i=(1):(EndRelays)
    if RawRelays{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawRelays{2}(i);
        Node_Name{2}(n) = RawRelays{6}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawRelays{3}(i);
    end
end
%From of Breakers
for i=(1):(EndBreakers)
    if RawBreakers{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawBreakers{2}(i);
        Node_Name{2}(n) = RawBreakers{6}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawBreakers{3}(i);
    end
end
%From of Fuses
for i=(1):(EndFuses)
    if RawFuses{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawFuses{2}(i);
        Node_Name{2}(n) = RawFuses{6}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawFuses{3}(i);
    end
end
%from of groundings
for i=(1):(EndGnds)
    if RawGnds{4}(i)
        n=n+1;
        Node_Name{1}(n) = RawGnds{2}(i);
        Node_Name{2}(n) = RawGnds{5}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawGnds{3}(i);
    end
end
%to of groundings
for i=(1):(EndGnds)
    if RawGnds{4}(i)
        n=n+1;
        Node_Name{1}(n) = {'NodeGND'};
        Node_Name{2}(n) = RawGnds{5}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawGnds{3}(i);
    end
end
%From of HVMVs
for i=(1):(EndHVMVs)
    if RawHVMVs{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawHVMVs{2}(i);
        Node_Name{2}(n) = RawHVMVs{6}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawHVMVs{3}(i);
    end
end
%from of regultors
for i=(1):(EndRegs)
    if RawReg{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawReg{2}(i);
        Node_Name{2}(n) = RawReg{6}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawReg{3}(i);
    end
end
%to of regultors
for i=(1):(EndRegs)
    if RawReg{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawReg{4}(i);
        Node_Name{2}(n) = RawReg{6}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawReg{3}(i);
    end
end
%from of trans
for i=(1):(EndTrans)
    n=n+1;
    Node_Name{1}(n) = RawTrans{2}(i);
    Node_Name{2}(n) = RawTrans{5}(i)*1000;
    Node_Name{3}(n) = 1;
    Node_Name{4}(n) = RawTrans{3}(i);
end
%to of trans
for i=(1):(EndTrans)
    n=n+1;
    Node_Name{1}(n) = RawTrans{4}(i);
    Node_Name{2}(n) = RawTrans{6}(i)*1000;
    Node_Name{3}(n) = 2;
    Node_Name{4}(n) = RawTrans{3}(i);
end
%loads
for i=(1):(EndLoads)
    if RawLoads{7}(i)
        n=n+1;
        Node_Name{1}(n) = RawLoads{2}(i);
        Node_Name{2}(n) = RawLoads{5}(i);
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = RawLoads{4}(i);
    end
end
% Go through 'to' node list
for i=(1):(EndLines)
    if RawLines{9}(i)
        n=n+1;
        Node_Name{1}(n) = RawLines{4}(i);
        Node_Name{2}(n) = RawLines{10}(i);
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = RawLines{3}(i);
    end
end
% Go through 'to' node list
for i=(1):(EndUtil)
    n=n+1;
    Node_Name{1}(n) = RawUtil{2}(i);
    Node_Name{2}(n) = RawUtil{4}(i);
    Node_Name{3}(n) = 2;
    Node_Name{4}(n) = RawUtil{3}(i);
end
%Switches placed at bottom to ensure that
%placeholder voltage is not prefered in the subsequent sort

%to of switches
for i=(1):(EndSwitches)
    sw2 = RawSwitches{3}(i);
    if (isempty(sw2{1})&& RawSwitches{5}(i))||~isempty(sw2{1})
        n=n+1;
        Node_Name{1}(n) = RawSwitches{4}(i);
        Node_Name{2}(n) = 999999999;
        Node_Name{3}(n) = 2;
        Node_Name{4}(n) = {'ABC'};
    end
end
%from of switches1
for i=(1):(EndSwitches)
    if RawSwitches{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawSwitches{2}(i);
        Node_Name{2}(n) = 999999999;
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = {'ABC'};
    end
end
%from of switches2
for i=(1):(EndSwitches)
    if ~RawSwitches{5}(i)
        n=n+1;
        Node_Name{1}(n) = RawSwitches{3}(i);
        Node_Name{2}(n) = 999999999;
        Node_Name{3}(n) = 1;
        Node_Name{4}(n) = {'ABC'};
    end
end

%% sort the list of nodes to find all node which appear exactly once
%sort the list of nodes to quickly delete duplicates
[Node_Name{1}, Nodes_order] = sort(Node_Name{1,1});
Node_Name{2} = Node_Name{1,2}(:,Nodes_order);
Node_Name{3} = Node_Name{1,3}(:,Nodes_order);
Node_Name{4} = Node_Name{1,4}(:,Nodes_order);
longPhase = 0;
k = 1;
id = [];
listlen = length(Node_Name{1});
for i = 1:listlen
    curr = Node_Name{1}(i);
    curlen = length(char(Node_Name{4}(i)));
    if curlen>longPhase
        longPhase =  curlen;
        k = i;
    end
    if i~=listlen
        next = Node_Name{1}(i+1);
        if ~strcmp(next,curr)
            id = [id,k];
            longPhase = 0;
        end
    else
        id = [id,k];
    end
end
Node_Name{1} = Node_Name{1,1}(:,id);
Node_Name{2} = Node_Name{1,2}(:,id);
Node_Name{3} = Node_Name{1,3}(:,id);
Node_Name{4} = Node_Name{1,4}(:,id);
Node_Name{1} = transpose(Node_Name{1});
Node_Name{2} = transpose(Node_Name{2});
Node_Name{3} = transpose(Node_Name{3});
Node_Name{4} = transpose(Node_Name{4});


% Print Nodes, but override all of the capacitor nodes to be three phase
swing = 0;
Node_Buser{1} = RawUtil{2};
Node_Buser{2} = RawUtil{4};
for i=1:length(Node_Name{1})
    phase = char(Node_Name{4}(i));
    if ~isempty(char(Node_Name{1}(i)))
        if sum(strcmp(Node_Buser{1},[char(Node_Name{1}(i))]))
            fprintf(fid,'object node {\n');
            fprintf(fid,'     phases %s;\n',phase);
            fprintf(fid,'     bustype SWING;\n');
            fprintf(fid,'     name "%s";\n',char(Node_Name{1}(i)));
            fprintf(fid,'     nominal_voltage %5.0f;\n',Node_Name{2}(i));
            fprintf(fid,'}\n\n');
            swing = 1;
        elseif strcmp(char(Node_Name{1}(i)),'NodeGND')
            fprintf(fid,'object node {\n');
            fprintf(fid,'     phases ABCN;\n');
            %             fprintf(fid,'     bustype PV;\n');
            fprintf(fid,'     name "NodeGND";\n');
            fprintf(fid,'     nominal_voltage 34500;\n');
            fprintf(fid,'}\n\n');
        else
            fprintf(fid,'object node {\n');
            fprintf(fid,'     phases %s;\n',phase);
            fprintf(fid,'     name "%s";\n',char(Node_Name{1}(i)));
            fprintf(fid,'     nominal_voltage %5.0f;\n',Node_Name{2}(i));
            fprintf(fid,'}\n\n');
        end
    end
end
if ~swing
  disp('Warning: No SWING buses were assigned, this may cause issues during simulation.');  
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Breakers
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Breakers.\n');
disp('Printing Breakers...');
for i = 1:EndBreakers
    if  RawBreakers{5}(i)== 1
        fprintf(fid,'object fuse {\n');
        fprintf(fid,'     name "%s";\n',char(RawBreakers{1}(i)));
        fprintf(fid,'     phases ABC;\n');
        fprintf(fid,'     from "%s";\n',char(RawBreakers{2}(i)));
        fprintf(fid,'     to "%s";\n',char(RawBreakers{4}(i)));
        fprintf(fid,'     current_limit %.1f A;\n',RawBreakers{7}(i));
        %         fprintf(fid,'     current_limit 9999.0 A;\n');
        fprintf(fid,'     mean_replacement_time 3600.0;\n');
        fprintf(fid,'     repair_dist_type NONE;\n');
        fprintf(fid,'}\n\n');
    end
end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Fuses
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Fuses.\n');
disp('Printing Fuses...');
for i = 1:EndFuses
    if  RawFuses{5}(i)== 1
        fprintf(fid,'object fuse {\n');
        fprintf(fid,'     name "%s";\n',char(RawFuses{1}(i)));
        fprintf(fid,'     phases ABCN;\n');
        fprintf(fid,'     from "%s";\n',char(RawFuses{2}(i)));
        fprintf(fid,'     to "%s";\n',char(RawFuses{4}(i)));
        fprintf(fid,'     current_limit %.1f A;\n',RawFuses{7}(i));
        %         fprintf(fid,'     current_limit 9999.0 A;\n');
        fprintf(fid,'     mean_replacement_time 3600.0;\n');
        fprintf(fid,'     repair_dist_type NONE;\n');
        fprintf(fid,'}\n\n');
    end
end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Generators
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%this implimentation may result in generators becoming swing buses and
%still supplying power
fprintf(fid,'// Generators.\n');
disp('Printing Generators...');
for i = 1:EndGenes
    if  RawGenes{5}(i)== 1
        fprintf(fid,'object load {\n');
        fprintf(fid,'     name "%s";\n',char(RawGenes{1}(i)));
        fprintf(fid,'     parent "%s";\n',char(RawGenes{2}(i)));
        fprintf(fid,'     nominal_voltage %7.0f;\n',RawGenes{3}(i));
        fprintf(fid,'     phases ABCN;\n');
        
        reload = RawGenes{5}(i).*1000/3;
        imload = RawGenes{5}(i)*tan(acos(RawGenes{6}(i))).*1000/3;
        
        fprintf(fid,'     constant_power_A -%.1f-%.1fj;\n',reload,imload);
        fprintf(fid,'     constant_power_B -%.1f-%.1fj;\n',reload,imload);
        fprintf(fid,'     constant_power_C -%.1f-%.1fj;\n',reload,imload);
        fprintf(fid,'}\n\n');
    end
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create HVMVs
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// HVMVs.\n');
disp('Printing HVMVs...');
for i = 1:EndHVMVs
    if  RawHVMVs{5}(i)== 1
        fprintf(fid,'object fuse {\n');
        fprintf(fid,'     name "%s";\n',char(RawHVMVs{1}(i)));
        fprintf(fid,'     phases ABCN;\n');
        fprintf(fid,'     from "%s";\n',char(RawHVMVs{2}(i)));
        fprintf(fid,'     to "%s";\n',char(RawHVMVs{4}(i)));
        fprintf(fid,'     current_limit 9999.0 A;\n');
        fprintf(fid,'     mean_replacement_time 3600.0;\n');
        fprintf(fid,'     repair_dist_type NONE;\n');
        fprintf(fid,'}\n\n');
    end
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Motor Controllers
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Motor Controllers.\n');
disp('Printing MTRCTRLs...');
fprintf(fid,'object line_configuration {\n');
fprintf(fid,'     name MotorControllerConfig;\n');
fprintf(fid,'     z11 0.03+0.01j;\n');
fprintf(fid,'     z12 0.01+0.01j;\n');
fprintf(fid,'     z13 0.01+0.01j;\n');
fprintf(fid,'     z21 0.01+0.01j;\n');
fprintf(fid,'     z22 0.03+0.01j;\n');
fprintf(fid,'     z23 0.01+0.01j;\n');
fprintf(fid,'     z31 0.01+0.01j;\n');
fprintf(fid,'     z32 0.01+0.01j;\n');
fprintf(fid,'     z33 0.03+0.01j;\n');
fprintf(fid,'}\n\n');
for i = 1:EndMTRCTRLs
    fprintf(fid,'object overhead_line {\n');
    fprintf(fid,'     name "%s";\n',char(RawMTRCTRLs{1}(i)));
    fprintf(fid,'     phases ABCN;\n');
    fprintf(fid,'     from "%s";\n',char(RawMTRCTRLs{2}(i)));
    fprintf(fid,'     to "%s";\n',char(RawMTRCTRLs{3}(i)));
    fprintf(fid,'     length 1;\n');
    fprintf(fid,'     configuration MotorControllerConfig;\n');
    fprintf(fid,'}\n\n');
    P = RawMTRCTRLs{5}(i);
    eff = 1/RawMTRCTRLs{7}(i);
    real = (eff-1)*P/3;
    imag = eff*RawMTRCTRLs{6}(i)*P/3;
    fprintf(fid,'object load {\n');
    fprintf(fid,'     name "%sLoad";\n',char(RawMTRCTRLs{1}(i)));
    fprintf(fid,'     parent "%s";\n',char(RawMTRCTRLs{3}(i)));
    fprintf(fid,'     nominal_voltage %7.0f;\n',RawMTRCTRLs{4}(i));
    fprintf(fid,'     phases ABCN;\n');
    fprintf(fid,'     constant_power_A %.1f+%.1fj;\n',real,imag);
    fprintf(fid,'     constant_power_B %.1f+%.1fj;\n',real,imag);
    fprintf(fid,'     constant_power_C %.1f+%.1fj;\n',real,imag);
    fprintf(fid,'}\n\n');
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create MTRPros
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// MTRPros.\n');
disp('Printing MTRPros...');
for i = 1:EndMTRPros
    if  RawMTRPros{5}(i)== 1
        fprintf(fid,'object fuse {\n');
        fprintf(fid,'     name "%s";\n',char(RawMTRPros{1}(i)));
        fprintf(fid,'     phases ABCN;\n');
        fprintf(fid,'     from "%s";\n',char(RawMTRPros{2}(i)));
        fprintf(fid,'     to "%s";\n',char(RawMTRPros{4}(i)));
        fprintf(fid,'     current_limit 9999.0 A;\n');
        fprintf(fid,'     mean_replacement_time 3600.0;\n');
        fprintf(fid,'     repair_dist_type NONE;\n');
        fprintf(fid,'}\n\n');
    end
end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Switches
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Switches\n');
disp('Printing Switches...');
for i = 1:EndSwitches
    if RawSwitches{5}(i) == 1
        position1 = 'CLOSED';
        position2 = 'OPEN';
    else
        position1 = 'OPEN';
        position2 = 'CLOSED';
    end
    if ~isempty(char(RawSwitches{2}(i))) && strcmp(position1,'CLOSED')
        fprintf(fid,'object switch {\n');
        fprintf(fid,'     name "%ssw1";\n',char(RawSwitches{1}(i)));
        fprintf(fid,'     from "%s";\n',char(RawSwitches{2}(i)));
        fprintf(fid,'     to "%s";\n',char(RawSwitches{4}(i)));
        fprintf(fid,'     phase_A_state %s;\n',position1);
        fprintf(fid,'     phase_B_state %s;\n',position1);
        fprintf(fid,'     phase_C_state %s;\n',position1);
        fprintf(fid,'     phases ABCN;\n');
        fprintf(fid,'}\n\n');
    end
    if ~isempty(char(RawSwitches{3}(i))) && strcmp(position2,'CLOSED')
        fprintf(fid,'object switch {\n');
        fprintf(fid,'     name "%ssw2";\n',char(RawSwitches{1}(i)));
        fprintf(fid,'     from "%s";\n',char(RawSwitches{3}(i)));
        fprintf(fid,'     to "%s";\n',char(RawSwitches{4}(i)));
        fprintf(fid,'     phase_A_state %s;\n',position2);
        fprintf(fid,'     phase_B_state %s;\n',position2);
        fprintf(fid,'     phase_C_state %s;\n',position2);
        fprintf(fid,'     phases ABCN;\n');
        fprintf(fid,'}\n\n');
    end
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Relays
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Relays.\n');
disp('Printing Relays...');
for i = 1:EndRelays
    if  RawRelays{5}(i)== 1
        fprintf(fid,'object recloser {\n');
        fprintf(fid,'     name "%s";\n',char(RawRelays{1}(i)));
        fprintf(fid,'     phases ABCN;\n');
        fprintf(fid,'     from "%s";\n',char(RawRelays{2}(i)));
        fprintf(fid,'     to "%s";\n',char(RawRelays{4}(i)));
        fprintf(fid,'     retry_time 1.0s;\n');
        fprintf(fid,'     max_number_of_tries 5;\n');
        fprintf(fid,'}\n\n');
    end
end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Line and Conductor Configurations
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
fprintf(fid,'// Overhead Line Conductors and configurations.\n');
disp('Printing lines and conductors...');
% Print the conductors that are needed
for i = 1:EndLines
    if RawLines{9}(i)&& isempty(char(RawLines{17}(i)))
        if contains(char(RawLines{14}(i)),{'Buri','Duct','UG','nder','Conduit'})
            OD = RawLines{15}(i);
            if isnan(RawLines{15}(i))
                OD = sqrt(RawLines{8}(i))/600;
            end
            shieldOD = RawLines{16}(i);
            if isnan(RawLines{16}(i))
                shieldOD = sqrt(RawLines{8}(i))/1000;
            end
            fprintf(fid,'object underground_line_conductor {\n');
            fprintf(fid,'     name %s_cond;\n',char(RawLines{1}(i)));
            fprintf(fid,'     outer_diameter %1.6f in;\n',OD);
            fprintf(fid,'     conductor_gmr %1.6f ft;\n',sqrt(RawLines{8}(i))/1000/12*0.42);
            fprintf(fid,'     conductor_diameter %.6f in;\n',shieldOD);
            fprintf(fid,'     conductor_resistance %1.6f Ohm/mile;\n',RawLines{11}(i)*5280/1000);
            fprintf(fid,'     insulation_relative_permitivitty 2.3;\n');
            fprintf(fid,'     shield_diameter %.2f in;\n', shieldOD*2);
            fprintf(fid,'     shield_thickness 0.08 in;\n');
            fprintf(fid,'}\n\n');
            fprintf(fid,'object line_configuration {\n');
            fprintf(fid,'     name %sConfig;\n',char(RawLines{1}(i)));
            fprintf(fid,'     conductor_A "%s_cond";\n',char(RawLines{1}(i)));
            fprintf(fid,'     conductor_B "%s_cond";\n',char(RawLines{1}(i)));
            fprintf(fid,'     conductor_C "%s_cond";\n',char(RawLines{1}(i)));
            fprintf(fid,'     spacing %sSpacing;\n',char(RawLines{1}(i)));
            fprintf(fid,'}\n\n');
            spacing = OD*0.0254;
            fprintf(fid,'object line_spacing {\n');
            fprintf(fid,'     name %sSpacing;\n',char(RawLines{1}(i)));
            fprintf(fid,'     distance_AB %.3fm;\n',spacing);
            fprintf(fid,'     distance_AC %.3fm;\n',spacing);
            fprintf(fid,'     distance_BC %.3fm;\n',spacing);
            fprintf(fid,'}\n\n');
        else
            fprintf(fid,'object overhead_line_conductor {\n');
            fprintf(fid,'     name %s_cond;\n',char(RawLines{1}(i)));
            fprintf(fid,'     geometric_mean_radius %1.6fcm;\n',sqrt(RawLines{8}(i))/2/1000*2.54);
            fprintf(fid,'     resistance %1.6f Ohm/mile;\n',RawLines{11}(i)*5280/1000);
            fprintf(fid,'     rating.summer.emergency %.0f A;\n',RawLines{12}(i));
            fprintf(fid,'     rating.summer.continuous %.0f A;\n',RawLines{12}(i));
            fprintf(fid,'     rating.winter.emergency %.0f A;\n',RawLines{12}(i));
            fprintf(fid,'     rating.winter.continuous %.0f A;\n',RawLines{12}(i));
            fprintf(fid,'}\n\n');
            fprintf(fid,'object line_configuration {\n');
            fprintf(fid,'     name %sConfig;\n',char(RawLines{1}(i)));
            fprintf(fid,'     conductor_A "%s_cond";\n',char(RawLines{1}(i)));
            fprintf(fid,'     conductor_B "%s_cond";\n',char(RawLines{1}(i)));
            fprintf(fid,'     conductor_C "%s_cond";\n',char(RawLines{1}(i)));
            fprintf(fid,'     conductor_N "%s_cond";\n',char(RawLines{1}(i)));
            fprintf(fid,'     spacing ThreePhase1;\n');
            fprintf(fid,'}\n\n');
        end
    end
end

for i = 1:EndCondConfig
    type = char(RawCondConfig{2}(i));
    if strcmp(type,'CN')||strcmp(type,'TS')
        fprintf(fid,'object underground_line_conductor {\n');
        fprintf(fid,'     name %s;\n',char(RawCondConfig{1}(i)));
        if ~isnan(RawCondConfig{7}(i))
            fprintf(fid,'     outer_diameter %1.6f in;\n',RawCondConfig{7}(i));
        end
        if ~isnan(RawCondConfig{10}(i))
            fprintf(fid,'     conductor_gmr %1.6f ft;\n',RawCondConfig{10}(i));
        end
        fprintf(fid,'     conductor_diameter %.3f in;\n',RawCondConfig{8}(i));
        fprintf(fid,'     conductor_resistance %1.6f Ohm/mile;\n',RawCondConfig{5}(i)*5280/1000);
        if strcmp(type,'CN')
            if ~isnan(RawCondConfig{11}(i))
                fprintf(fid,'     neutral_gmr %1.6f ft;\n',RawCondConfig{11}(i));
                if ~isnan(RawCondConfig{9}(i))
                    fprintf(fid,'     neutral_diameter %.3f in;\n',RawCondConfig{9}(i));
                end
            else
                if ~isnan(RawCondConfig{9}(i))
                    fprintf(fid,'     neutral_diameter %.3f in;\n',RawCondConfig{9}(i));
                else
                    disp(['"Warning: There is insufficient information to define the concentric neutral of conductor "',char(RawCondConfig{1}(i)),'".']);
                end
            end
            if ~isnan(RawCondConfig{6}(i))
                fprintf(fid,'     neutral_resistance %1.6f Ohm/mile;\n',RawCondConfig{6}(i)*5280/1000);
            end
            fprintf(fid,'     neutral_strands %1.f;\n',RawCondConfig{4}(i));
        end
        fprintf(fid,'     insulation_relative_permitivitty 2.3;\n');
%         if ~isnan(RawCondConfig{13}(i))
%             fprintf(fid,'     shield_diameter %.2f in;\n', RawCondConfig{13}(i));
%             fprintf(fid,'     shield_thickness %.1f in;\n',RawCondConfig{14}(i));
%         end
            fprintf(fid,'     shield_gmr 0.0;\n');
            fprintf(fid,'     shield_resistance 0.0;\n');
        fprintf(fid,'}\n\n');
    else
        fprintf(fid,'object overhead_line_conductor {\n');
        fprintf(fid,'     name %s;\n',char(RawCondConfig{1}(i)));
        fprintf(fid,'     geometric_mean_radius %1.6fin;\n',RawCondConfig{10}(i));
        fprintf(fid,'     diameter %1.6fin;\n',RawCondConfig{8}(i));
        fprintf(fid,'     resistance %1.6f Ohm/mile;\n',RawCondConfig{5}(i)*5280/1000);
        fprintf(fid,'     rating.summer.emergency %.0f A;\n',RawCondConfig{12}(i));
        fprintf(fid,'     rating.summer.continuous %.0f A;\n',RawCondConfig{12}(i));
        fprintf(fid,'     rating.winter.emergency %.0f A;\n',RawCondConfig{12}(i));
        fprintf(fid,'     rating.winter.continuous %.0f A;\n',RawCondConfig{12}(i));
        fprintf(fid,'}\n\n');
    end
end

for i = 1:EndConfig
    fprintf(fid,'object line_configuration {\n');
    fprintf(fid,'     name %s;\n',char(RawConfig{1}(i)));
    temp = RawConfig{2}(i);
    if ~isempty(temp{1})
    fprintf(fid,'     conductor_A "%s";\n',char(RawConfig{2}(i)));
    end
    temp = RawConfig{3}(i);
    if ~isempty(temp{1})
    fprintf(fid,'     conductor_B "%s";\n',char(RawConfig{3}(i)));
    end
    temp = RawConfig{4}(i);
    if ~isempty(temp{1})
    fprintf(fid,'     conductor_C "%s";\n',char(RawConfig{4}(i)));
    end
    temp = RawConfig{5}(i);
    if ~isempty(temp{1})
    fprintf(fid,'     conductor_N "%s";\n',char(RawConfig{5}(i)));
    end
    fprintf(fid,'     spacing %s;\n',char(RawConfig{6}(i)));
    fprintf(fid,'}\n\n');
end


% Create line spacings
for i = 1:EndSpacing
    fprintf(fid,'object line_spacing {\n');
    fprintf(fid,'     name %s;\n',char(RawSpacing{1}(i)));
    if ~isnan(RawSpacing{2}(i))
    fprintf(fid,'     distance_AB %.4fft;\n',RawSpacing{2}(i));
    end
    if ~isnan(RawSpacing{3}(i))
    fprintf(fid,'     distance_AC %.4fft;\n',RawSpacing{3}(i));
    end
    if ~isnan(RawSpacing{4}(i))
    fprintf(fid,'     distance_BC %.4fft;\n',RawSpacing{4}(i));
    end
    if ~isnan(RawSpacing{5}(i))
    fprintf(fid,'     distance_AN %.4fft;\n',RawSpacing{5}(i));
    end
    if ~isnan(RawSpacing{6}(i))
    fprintf(fid,'     distance_BN %.4fft;\n',RawSpacing{6}(i));
    end
    if ~isnan(RawSpacing{7}(i))
    fprintf(fid,'     distance_CN %.4fft;\n',RawSpacing{7}(i));
    end
    if ~isnan(RawSpacing{8}(i))
    fprintf(fid,'     distance_AE %.4fft;\n',RawSpacing{8}(i));
    end
    if ~isnan(RawSpacing{9}(i))
    fprintf(fid,'     distance_BE %.4fft;\n',RawSpacing{9}(i));
    end
    if ~isnan(RawSpacing{10}(i))
    fprintf(fid,'     distance_CE %.4fft;\n',RawSpacing{10}(i));
    end
    if ~isnan(RawSpacing{11}(i))
    fprintf(fid,'     distance_NE %.4fft;\n',RawSpacing{11}(i));
    end
    fprintf(fid,'}\n\n');
end
fprintf(fid,'object line_spacing {\n');
fprintf(fid,'     name ThreePhase1;\n');
fprintf(fid,'     distance_AB 0.97584m;\n');
fprintf(fid,'     distance_AC 1.2192m;\n');
fprintf(fid,'     distance_BC 0.762m;\n');
fprintf(fid,'     distance_BN 2.1336m;\n');
fprintf(fid,'     distance_AN 1.70388m;\n');
fprintf(fid,'     distance_CN 1.5911m;\n');
fprintf(fid,'}\n\n');

fprintf(fid,'object line_spacing {\n');
fprintf(fid,'     name UndergroundSpacing1;\n');
fprintf(fid,'     distance_AB 0.03m;\n');
fprintf(fid,'     distance_AC 0.03m;\n');
fprintf(fid,'     distance_BC 0.03m;\n');
fprintf(fid,'}\n\n');

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create line objects
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

fprintf(fid,'// Overhead/Underground Lines\n\n');

for i=1:EndLines
    if ~isempty(RawConfig)
        [conf{1},I] = sort(RawConfig{1});
        for j = 2:5
            conf{j} = RawConfig{j}(I);
        end
    end
    if RawLines{9}(i)
        config = RawLines{1}(i);
        add = 'Config';
        phases = char(RawLines{3}(i));
        if isempty(char(RawLines{14}(i)))
            disp(['Warning: Installation type is not defined for conductor "',config,'".']);
        end
        conff = RawLines{17}(i);
        if ~isempty(conff{1})
            config = RawLines{17}(i);
            add = '';
            phases = '';
            id = binsearch(RawConfig{1},config);
            check = conf{2}(id)
            if ~isempty(check{1})
                phases = 'A';
            end
            check = conf{3}(id);
            if ~isempty(check{1})
                phases = [phases,'B'];
            end
            check = conf{4}(id);
            if ~isempty(check{1})
                phases = [phases,'C'];
            end
            check = conf{5}(id);
            if ~isempty(check{1})
                phases = [phases,'N'];
            end
        end
        if LengthLines(i)==0 || isempty(LengthLines(i))
            LengthLines(i) = 1;
            RawLines{7}(i) = 'ft';
        end
        if contains(char(RawLines{14}(i)),{'Buri','Duct','UG','nder','Conduit'})
            fprintf(fid,'object underground_line {\n');
            fprintf(fid,'     phases %s;\n',phases);
            fprintf(fid,'     name "%s";\n',char(RawLines{1}(i)));
            fprintf(fid,'     from "%s";\n',char(RawLines{2}(i)));
            fprintf(fid,'     to "%s";\n',char(RawLines{4}(i)));
            fprintf(fid,'     length %f%s;\n',LengthLines(i),char(RawLines{7}(i)));
            fprintf(fid,'     configuration "%s%s";\n',char(config{1}),add);
            fprintf(fid,'}\n\n');
        else
            fprintf(fid,'object overhead_line {\n');
            fprintf(fid,'     phases %s;\n',phases);
            fprintf(fid,'     name "%s";\n',char(RawLines{1}(i)));
            fprintf(fid,'     from "%s";\n',char(RawLines{2}(i)));
            fprintf(fid,'     to "%s";\n',char(RawLines{4}(i)));
            fprintf(fid,'     length %f%s;\n',LengthLines(i),char(RawLines{7}(i)));
            fprintf(fid,'     configuration "%s%s";\n',char(config{1}),add);
            fprintf(fid,'}\n\n');
        end
    end
end
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Triplex-Line objects
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

fprintf(fid,'// Triplex Lines\n\n');
disp('Printing triplex lines...');
for i=1:EndTripLines
    fprintf(fid,'object triplex_line {\n');
    fprintf(fid,'     name "%s";\n',char(RawTripLines{1}(i)));
    Tp = char(RawTripLines{4}(i));
    Tphase = Tp(10);
    fprintf(fid,'     phases %sS;\n',Tphase);
    fprintf(fid,'     from "%s";\n',char(RawTripLines{2}(i)));
    fprintf(fid,'     to "%s";\n',char(RawTripLines{4}(i)));
    if (strcmp(houses,'y')~= 0)
        fprintf(fid,'     length %.1fft;\n',25-20*rand(1));
    else
        fprintf(fid,'     length %2.0fft;\n',RawTripLines{7}(i));
    end
    fprintf(fid,'     configuration "%s";\n',char(RawTripLines{6}(i)));
    fprintf(fid,'}\n\n');
end

%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Cleanup GLM
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
disp('Cleaning up file...');
GLDCleanup(open_name);
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Create Recorders
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
disp('Creating recorder file...');
GLDrecorder(dir,fileName,recorder);
%%
fclose('all');
disp(['File generation completed. Elapsed Time: ' num2str(toc)]);