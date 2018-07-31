%% Take all CSVs from a folder and compile them into a single sheet and graph.
%%
clear all
dirFile = 'graphDir.txt';
keepCurrentSettings = 'No';
if ~exist(dirFile)
    fid = fopen(dirFile,'w');
    fclose(fid);
end
dirFileid = fopen(dirFile,'r');
initDir1 = fgetl(dirFileid);
PFdir = char(fgetl(dirFileid));
PFname = fgetl(dirFileid);
fclose(dirFileid);
if ~ischar(initDir1)
    initDir1 = pwd;
end
dir = char(initDir1);
keepCurrentSettings = questdlg('Use previous folder?',...
    'Confirmation',...
    'Yes','No','Yes');
switch keepCurrentSettings
    case 'Yes'
    case 'No'
end

if strcmp(keepCurrentSettings,'No')
    dir = uigetdir(initDir1,'Select recorder directory');
    dirFileid = fopen(dirFile,'w');
    fprintf(dirFileid,'%s',dir);
    fprintf(dirFileid,'\n');
    [PFname,PFdir,filter] = uigetfile({'*.csv;*.xls;*xlsx','All Tabular Files';'*.*','All Files' },'Select your known-power file',PFdir);
    fprintf(dirFileid,'%s',PFdir);
    fprintf(dirFileid,'\n');
    fprintf(dirFileid,'%s',PFname);
    fprintf(dirFileid,'\r\n');
    fclose(dirFileid);
end
if strcmp(keepCurrentSettings,'Yes') &&  strcmp(initDir1,pwd)
    disp('Error: There are no files or directories given, you must select new settings.');
    return;
end
%%
tic
file = ls(dir);
numFiles = numrows(file);
k = 3;
for i = 3:1:numFiles
    currfile = strtrim(file(i,:));
    if contains(currfile,'.csv')
        j = k-2;
        file2find = [dir,'\',currfile];
        P = csvread(file2find,10,1);
        name{j,1} = regexprep(currfile,{'slash'},{'/'});
        try
            sign = mean(P(:,1));
            RMSre{j,1} = sign/(sign)*rms(P(:,1));
        catch
        end
        
        try
            sign = mean(P(:,2));
            RMSim{j,1} = sign/(sign)*rms(P(:,2));
        catch
        end
        
        try
            RMSVre{j,1} = rms(P(:,3));
        catch
        end
        
        try
            RMSVim{j,1} = rms(P(:,4));
        catch
        end
        
        try
            RMSV{j,1} = sqrt(RMSVre{j,1}^2+RMSVim{j,1}^2);
        catch
        end
        k = k+1;
    end
end
%%
file2find = [PFdir,PFname];
if exist(file2find)
    name = regexprep(name,{'.csv'},{''}); %no need to sort, done by file system
    PKfile = fopen(file2find);
    PKheader = textscan(PKfile,'%s %s %s %s %s %s %s %s',1,'Delimiter',',');
    PK = textscan(PKfile,'%s %n %n %n %n %n %n %n ','Delimiter',',');
    len = numrows(PK);
    namek = PK{1};
    namek = regexprep(namek,{':','&',',',' '},{'CNXN','and','P',''});
    [namek,I] = sort(namek);
    RMSrek = PK{6}(I);
    RMSimk = PK{5}(I);
    Vk = PK{8}(I);
    j = 1;
    Error = {};
    for i = 1:length(namek)
        nameKnown = namek(i);
        id = binsearch(name,nameKnown);
        if id ~= 0
            entryre = calcerror(RMSrek(i,1),RMSre{id,1}/1000);
            entryim = calcerror(RMSimk(i,1),RMSim{id,1}/1000);
            try
                voltage = calcerror(Vk(i,1),RMSV{id,1});
            catch
                voltage = 0;
            end
            if entryre == inf %uncomment if these unloaded components should be graphed
                %                 Error{j,1} = 0;
                %                 Error{j,2} = 0;
                %                 Error{j,3} = 0;
                %                 namee(j,1) = name(id);
                %                 j = j+1;
            else
                %elseif entryre <50 && entryre>-50 % restrict the output to +-50%
                Error{j,1} = entryre;
                Error{j,2} = entryim;
                Error{j,3} = voltage;
                namee(j,1) = name(id);
                j = j+1;
            end
        end
    end
    bar(cell2mat(Error(:,1:3)))% 1 = real, 2 = imag, 3 = V
    title('Power Flow Percent Error')
    xticks(1:length(namee))
    xticklabels(namee)
    xtickangle(60)
    ylabel('Error (%)')
    legend('Real','Imaginary','Voltage')
end
disp(['Elapsed Time: ',num2str(toc)])