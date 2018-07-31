function [num,txt,raw] = tryxlsread(file,sheetname,depth)
%UNTITLED2 Summary of this function goes here
%   Detailed explanation goes here
[num,txt,raw] = deal([]);
try
[num,txt,raw] = xlsread(file,sheetname,depth);
catch
    try
        [num,txt,raw] = xlsread(file,sheetname);
    catch
    disp(['Warning: There was no sheet called "',sheetname,'" in file "',file,'".']);
    end
end
end

