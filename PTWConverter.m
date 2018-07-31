%% Fix Excel Sheets for Case Western Reserve
% Programmer: Emily Fabian, NASA Pathways Intern
% Date: September 2017
% Purpose: NASA exports a CSV from Powertools. This program then converts 
% the file into one that is usable for Case Western University.
% I know that this script is very klunky and messy, but it works!

% How to run: 
% Step 1: Press the green arrow under the "Editor" tab that says "Run"
% Step 2: Follow instructions on pop-up boxes
% Step 3: Go get a coffee, this code takes some time to run
% Step 4: Profit

clear all; close all; clc; 

%% dialog box
% create a dialog box to input the name of the file to be converted
dirFile = 'PTWDir.txt';
if ~exist(dirFile)
    msgbox({'Something went wrong! Make sure you spelled the file name correctly and that the file is in the MATLAB folder.' ' ' 'Please restart the script.'},'Error','error');
    fid = fopen(dirFile,'w');
    fclose(fid);
end
fid = fopen('PTWDir.txt','r');
dir = fgetl(fid);
file = fgetl(fid);
dir2 = fgetl(fid);
fclose(fid);
if ~ischar(dir)
    dir = pwd;
end
if ~ischar(dir2)
    dir2 = pwd;
end
dir = char(dir);
keepCurrentSettings = questdlg('Use previous file and settings?',...
    'Confirmation',...
    'Yes','No','File only','Yes');
switch keepCurrentSettings
    case 'Yes'
    case 'No'
    case 'File only'
end
if strcmp(keepCurrentSettings,'No')
    [file,dir,filter] = uigetfile({'*.csv;*.xls;*xlsx','All Tabular Files';'*.*','All Files' },'Select CSV file to convert',dir);
    dirFile = fopen('PTWDir.txt','w');
    fprintf(dirFile,'%s',dir);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',file);
    fprintf(dirFile,'\r\n');
    fclose(dirFile);
end
if strcmp(keepCurrentSettings,'File only') || strcmp(keepCurrentSettings,'No')
    dir2 = uigetdir(dir2,'Select Output file directory');
    dirFile = fopen('PTWDir.txt','w');
    fprintf(dirFile,'%s',dir);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',file);
    fprintf(dirFile,'\r\n');
    fprintf(dirFile,'%s',dir2);
    fprintf(dirFile,'\r\n');
    fclose(dirFile);
end
if strcmp(dir,pwd)||~ischar(file)||isempty(file)
    disp('Error: There are no files or directories given, you must select new settings.');
    return;
end
dirFile = [dir,file];
file2 = regexprep(file,{'\.csv'},{'New.xlsx'});
dirFile2 = [dir2,[file2]];
fid2 = fopen(dirFile2,'w');

filename = file; %name of the original file
outFile = file2; %name of the new file

%% Read input file and Error Messages

try %Get the file data
    [inputData,inputHeaders,raw] = xlsread(dirFile); 
catch %Display an error message if something goes wrong
    msgbox({'Something went wrong! Make sure you spelled the file name correctly and that the file is in the MATLAB folder.' ' ' 'Please restart the script.'},'Error','error');
    return
end

character = char(outFile);

if length(character) > 4
    if character(length(character)-4:length(character)) ~= '.xlsx'
        msgbox({'Something went wrong! Make sure you include the file extension ".xlsx" in the output file name.' ' ' 'Please restart the script.'},'Error','error');
        return
    end
else
    msgbox({'Something went wrong! Make sure you include the file extension ".xlsx" in the output file name.' ' ' 'Please restart the script.'},'Error','error');
    return
end

msgbox('Thank you, this could take a few minutes.');

%% Make first tab

xlswrite(outFile,raw); %creates the new excel file, first tab is the original file

%% Get rid of unnecessary rows in file

% delete first 10 rows
n = 11; %rows in original excel file
m = 0;
p = 1; %rows in matrix A

while n < length(raw)
    while m == 0 %loop until you hit an empty cell
        if isempty(char(inputHeaders(n,1))) == 0 %check if there is an empty cell
            A(p,:) = string(raw(n,:)); %if not, add it to the array A
            n = n+1; 
            p = p+1;
        else
            m = 1; %if so, then stop making the array
        end
        if n >= length(raw)
            m = 1;
        end
    end
    if n >= length(raw)
        break
    end
    p = p+1;
    m = 0;
    n = n+1;
    while m == 0
        if isempty(char(inputHeaders(n,2))) == 0
            A(p,:) = string(raw(n-1,:));
            p = p+1;
            A(p,:) = string(raw(n,:));
            p = p+1;
            n = n+1;
            break
        else
            n = n+1;
        end
        if n >= length(raw)
            break
        end
    end
end

A(p,:) = string(raw(n,:));

%% Start all sections on first line of file

n = 1; %rows in A
N = 1; %rows in B
p = 0; %columns in A
P = 0; %columns in B
m = 0;
l = 2;

while l < length(A)
    while m == 0 %count row headers
       try
           if ismissing(A(l,p+1)) == 0
               p = p+1;
               P = P+1;
           else
               m = 1;
           end
       catch
           m = 1;
       end
    end

    m = 0;

    while m == 0 %put section into new matrix
        try
            if ismissing(A(n,1)) == 0
                B(N,1+P-p:P) = A(n,1:p);
                n = n+1;
                N = N+1;
            else
                m = 1;
            end
        catch
            m = 1;
        end
    end
    p = 0;
    l = 2+n;
    n = n+1;
    N = 1;
    m = 0;
end

%% Make header names better

header1 = B(1,:); %first row headers
header2 = B(2,:); %second row headers

% fix header1
n = 1;
while n < length(header1)
    if ismissing(header1(n)) == 0
        title = char(header1(n));
        title = title(8:length(title)-4);
        header1(n) = string(title);
        n = n+1;
    else
        n = n+1;
    end    
end

%fix header2
n = 1;
m = 1;

while n <= length(header2)
    if ismissing(header2(n)) == 0
        title = char(header2(n));
        if title(1:2) == '//'
            title = title(3:length(title));
        end
        title = title(2:length(title)-1);
        for m = 1:length(title)
            if title(m) == '/' %replace slases with dashes
                title(m) = '-';
                m = m+1;
            else
                m = m+1;
            end
        end
        m = 1;
        if length(title) > 28 %If a header is too long
            prompt = (['The header "' title '" is too long. New header (Max 28 characters):']);
            title = inputdlg(prompt, 'Input'); %command
        end
        header2(n) = string(title);
        n = n+1;
    else
        n = n+1;
    end    
end

B(1,:) = header1;
B(2,:) = header2;

%% Export as an XLSX file

n = 1; %rows, a is max
m = 1; %columns, b is max
[a,b] = size(B);

for n = 1:a
   for m = 1:b
       if ismissing(B(n,m)) == 1
           B(n,m) = '';
       end      
       m = m+1;
   end
   n = n+1;
end

xlswrite('newFile1.xlsx',B); %create a temperary excel file called newFile1.xlsx
                             %this file will be deleted later in the script
                             
%% Clean up the variables

clear A;
clear a;
clear b;
clear B;
clear character;
clear defaultAns;
clear dlg_title;
clear header1;
clear header2;
clear l;
clear m;
clear n;
clear N;
clear num_lines;
clear p;
clear P;

%% Read input file Part II

[inputData,inputHeaders] = xlsread(char('newFile1.xlsx')); 

%% Convert the file into a usable format of tabs

n = 3;
m = 0;
f = char(inputHeaders(1,1));
header1 = [char(f(1)),char(f(2)),char('_'),char(inputHeaders(2,1))]; %create first header name
A(1,1) = string(inputHeaders(1,1));
A(2,1) = string(inputHeaders(2,1));

while m == 0 %only loop while m is 0
    try
        if isempty(char(inputHeaders(n,1))) == 0 %check if there is an empty cell
            A(n,1) = string(inputHeaders(n,1)); %if not, add it to the array A
            n = n+1; 
        else
            m = 1; %if so, then stop making the array
        end
    catch
        n = n-1;
        m = 1;
    end
end

xlswrite(outFile,A,header1); %create a new excel file

n = length(inputHeaders);
m = 0;
l = 2; %row
p = 2; %column
testH = 0;
testD = 0;
firstHeader = [char(f(1)),char(f(2)),char('_')];
A(n,1) = '';

while m == 0 %only loop while m is 0
    try
        if isempty(char(inputHeaders(1,p))) == 0 %check if 1st row has a header
            f = char(inputHeaders(1,p));
            firstHeader = [char(f(1)),char(f(2)),char('_')];
            A(1,1) = string(inputHeaders(1,p));
        end
    catch
        break
    end
    if isempty(char(inputHeaders(2,p))) == 0 %if there is not an empty header
        while l < n
           if isempty(char(inputHeaders(l+1,p))) == 0 % not empty, testH = 1
              testH = 1;
           end
           if isnan(inputData(l-1,p-1)) == 0 %not empty
              testD = 1;
           end 
           l = l+1;
        end
        l = 2;
        if testH == 1 && testD == 0 %if the collumn has words
            while l <= n 
                if isempty(char(inputHeaders(1,p))) == 0
                    A(l,1) = string(inputHeaders(l,p)); %add it to the array A
                    l = l+1;
                else
                    A(l-1,1) = string(inputHeaders(l,p));
                    l = l+1;
                end
            end
            header = [char(firstHeader),char(inputHeaders(2,p))];
            % Add after the last sheet
            xlswrite(outFile,A,header);
            p = p+1;
            l = 2;
        elseif testD == 1 && testH == 0 %if the collumn has numbers
            while l < n 
                if isnan(inputData(l-1,p-1)) == 0 %if cell is not empty
                    A(l,1) = string(inputData(l-1,p-1)); %add it to the array A
                    l = l+1;
                else %if cell is empty
                    A(l,1) = '';
                    l = l+1;
                end
            end
            A(1,1) = inputHeaders(2,p);
            header = [char(firstHeader),char(inputHeaders(2,p))];
            % Add after the last sheet
            xlswrite(outFile,A,header);
            p = p+1;
            l = 2;
        elseif testD == 1 && testH == 1 %if the collumn has numbers and words
            while l < n 
                if isnan(inputData(l-1,p-1)) == 0
                    A(l,1) = string(inputData(l-1,p-1)); %add it to the array A
                    l = l+1;
                elseif isempty(char(inputHeaders(l+1,p))) == 0
                    A(l,1) = string(inputHeaders(l+1,p));
                    l = l+1;
                else
                    A(l,1) = '';
                    l = l+1;
                end
            end
            A(1,1) = inputHeaders(2,p);
            header = [char(firstHeader),char(inputHeaders(2,p))];
            % Add after the last sheet
            xlswrite(outFile,A,header);
            p = p+1;
            l = 2;
        else
            p = p+1;
            l = 2;
        end
        A(n,1) = '';
        testD = 0;
        testH = 0;
        clc;
    else
        m = 1;
    end
end

delete newFile1.xlsx;

msgbox({'Done :)' 'Your new file is called ' outFile});