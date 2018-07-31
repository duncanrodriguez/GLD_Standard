clear all;
[GLM,dir,filter] = uigetfile({'*.glm;*.GLM','All GLD files';'*.*','All Files' },'Select GLM file to search');
dirFile = fopen('seekDir.txt','w');
fprintf(dirFile,'%s',dir);
fprintf(dirFile,'\r\n');
fprintf(dirFile,'%s',GLM);
fprintf(dirFile,'\r\n');
fclose(dirFile);
if ~ischar(initDir1)
    initDir1 = pwd;
end
if ~ischar(initDir2)
    initDir2 = pwd;
end
fid = fopen([dir,GLM]);
i = 1;
n = 1;
while ~feof(fid)
    newline = fgetl(fid);
    if sum(strfind(newline,'from '))~=0
        connection{1}(n) = cellstr(newline);
        newline = fgetl(fid);
        connection{2}(n) = cellstr(newline);
        connection{3}(n) = i+n;
        n = n+1;
    end
    i=i+1;
end
fclose(fid);
connection{1} = [connection{1},connection{2}];
connection{1} = regexprep(connection{1},{'from','to',',''',' '},{'','','',''});
connection{2} = [connection{2},connection{1}];
connection{2} = regexprep(connection{2},{'from','to',',''',' '},{'','','',''});
connection{3} = [connection{3},connection{3}];
[connection{1},I] = sort(connection{1});
connection{2} = connection{1,2}(:,I);
connection{3} = connection{1,3}(:,I);
connection{1}=transpose(connection{1});
connection{2}=transpose(connection{2});
connection{3}=transpose(connection{3});
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%Find duplicates
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
k=1;
for i = 1:length(connection{1})-3
    if strcmp(connection{1}(i),connection{1}(i+1))==1
        j=i+1;
        while strcmp(connection{1}(i),connection{1}(j))==1
            if strcmp(connection{2}(i),connection{2}(j))==1
                DupeAdmit(k,1)=connection{3}(i);
                DupeAdmit(k,2)=connection{3}(j);
                k=k+1;
            end
            j=j+1;
        end
    end
end
if exist('DupeAdmit') == 1
    [DA(:,1),I] = unique(DupeAdmit(:,1));
    DA2 = DupeAdmit(:,2);
    DA(:,2) = DA2(I);
    num = numrows(DA);
    disp(['"' GLM '" has ',num2str(num) ,' duplicate admittances at the following line combinations:'])
    disp(flipud(DA))
    disp('Remove one object and its configuration per combination.')
    disp(['Combinations are displayed in reverse numarical order to aid in deletion (This retains the accuracy of subsequent combinations).'])
else
    disp(['There are no duplicate admittances or paralellized components in "' GLM '".'])
end